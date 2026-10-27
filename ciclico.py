"""INJUL - Plan de Monitoreo Cíclico y Novedades (módulo independiente).

Es un Blueprint de Flask con sus propias tablas (prefijo ``ciclico_``), sus propias
rutas (``/ciclico``) y su propia interfaz. No toca los reportes consolidados, el
estado de cámaras ni los tickets del sistema principal: sólo LEE el catálogo
oficial de cámaras (``CATALOGO_CAMARAS``) que le entrega ``main.py``.

El catálogo que recibe es el CATÁLOGO CÍCLICO OFICIAL (la cuota de 5 a 10 cámaras
clave por sede que un operador audita por turno), no el inventario completo.

Todo se guarda en PostgreSQL al instante, vinculado por fecha, y cada cambio queda
en ``ciclico_auditoria`` para no perder el rastro de nada.
"""

import logging
import re
from datetime import datetime
from io import BytesIO
from xml.sax.saxutils import escape

from flask import Blueprint, Response, jsonify, render_template_string, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    Flowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

logger = logging.getLogger("injul-cctv.ciclico")

bp = Blueprint("ciclico", __name__, url_prefix="/ciclico")

# Dependencias inyectadas por registrar() (así este archivo no importa main.py).
_cfg = {}

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
UBICACION_DEFECTO = "MARMOTECH - Red de Sucursales"
AREA_DEFECTO = "Monitoreo / Sistema de Videovigilancia"
RE_HORA = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")

# Campos editables del turno: nombre -> (tipo, longitud máxima)
CAMPOS_TURNO = {
    "dia_rotacion": ("txt", 40),
    "desde": ("hora", 5),
    "hasta": ("hora", 5),
    "ubicacion_sistema": ("txt", 150),
    "responsable": ("txt", 120),
    "area": ("txt", 150),
    "informe_novedades": ("txt", 10000),
}
MAX_NOVEDADES_SITIO = 4000

# Cuota operativa mínima por sede y turno: un ciclo de 5 cámaras. Las sedes de 10
# cuadros (San Cristóbal, DCM) tienen dos ciclos: 1-5 (parte superior) y 6-10 (inferior).
CUOTA_MINIMA = 5
CICLO = 5

# Filas del "RESUMEN INFORMATIVO DEL PROCESO" tal como aparecen en el PDF oficial:
# (clave en el catálogo cíclico | None si no tiene bloque, rótulo oficial, total de cámaras).
RESUMEN_OFICIAL = [
    ("Sucursal San Cristóbal", "SAN CRISTÓBAL", "52"),
    ("Marmotech DCM", "SANTO DOMINGO (DCM)", "51"),
    ("Marmotech Santiago 1", "SANTIAGO 1", "16"),
    ("Marmotech Santiago 2", "SANTIAGO 2", "16"),
    ("Marmotech Blue Mall", "BLUE MALL", "16"),
    ("Marmotech Outlet RP", "OUTLET RP", "16"),
    ("Marmotech Punta Cana", "PUNTA CANA", "17"),
    ("Marmotech Caribbean", "CARIBBEAN", "22"),
    (None, "BOCA CHICA", ""),
]
TOTAL_SEDE = {clave: total for clave, _, total in RESUMEN_OFICIAL if clave}


# ----------------------------------------------------------------------------
# Base de datos
# ----------------------------------------------------------------------------
def crear_tablas(cur):
    cur.execute("""
        CREATE TABLE IF NOT EXISTS ciclico_turnos (
            fecha DATE PRIMARY KEY,
            dia_rotacion VARCHAR(40) NOT NULL DEFAULT '',
            desde VARCHAR(5) NOT NULL DEFAULT '',
            hasta VARCHAR(5) NOT NULL DEFAULT '',
            ubicacion_sistema VARCHAR(150) NOT NULL DEFAULT '',
            responsable VARCHAR(120) NOT NULL DEFAULT '',
            area VARCHAR(150) NOT NULL DEFAULT '',
            informe_novedades TEXT NOT NULL DEFAULT '',
            creado_en TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            actualizado_en TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
        CREATE TABLE IF NOT EXISTS ciclico_sitios (
            fecha DATE NOT NULL,
            sucursal VARCHAR(100) NOT NULL,
            hora_inicio VARCHAR(5) NOT NULL DEFAULT '',
            hora_fin VARCHAR(5) NOT NULL DEFAULT '',
            novedades TEXT NOT NULL DEFAULT '',
            actualizado_en TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            PRIMARY KEY (fecha, sucursal)
        );
        CREATE TABLE IF NOT EXISTS ciclico_camaras (
            fecha DATE NOT NULL,
            sucursal VARCHAR(100) NOT NULL,
            orden INT NOT NULL,
            camara VARCHAR(150) NOT NULL,
            revisada BOOLEAN NOT NULL DEFAULT FALSE,
            hora_inicio VARCHAR(5) NOT NULL DEFAULT '',
            hora_fin VARCHAR(5) NOT NULL DEFAULT '',
            actualizado_en TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            PRIMARY KEY (fecha, sucursal, orden)
        );
        CREATE TABLE IF NOT EXISTS ciclico_auditoria (
            id BIGSERIAL PRIMARY KEY,
            cambiado_en TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            fecha DATE NOT NULL,
            sucursal VARCHAR(100) NOT NULL DEFAULT '',
            camara VARCHAR(150) NOT NULL DEFAULT '',
            campo VARCHAR(40) NOT NULL,
            valor_anterior TEXT,
            valor_nuevo TEXT,
            origen VARCHAR(60)
        );
        CREATE INDEX IF NOT EXISTS idx_ciclico_aud_fecha
            ON ciclico_auditoria (fecha, cambiado_en DESC);
    """)


def _upsert(cur, tabla, pk, extra_insert, campos, camara_log=""):
    """Crea la fila si no existe, compara y actualiza sólo lo que cambió.

    Cada cambio se registra en ciclico_auditoria. Devuelve los valores finales.
    ``tabla`` y los nombres de columnas vienen siempre de constantes del código.
    """
    cols = list(pk) + list(extra_insert)
    cur.execute(
        f"INSERT INTO {tabla} ({','.join(cols)}) VALUES ({','.join(['%s'] * len(cols))}) "
        "ON CONFLICT DO NOTHING",
        [*pk.values(), *extra_insert.values()],
    )
    if not campos:
        return {}
    donde = " AND ".join(f"{k}=%s" for k in pk)
    nombres = list(campos)
    cur.execute(
        f"SELECT {','.join(nombres)} FROM {tabla} WHERE {donde} FOR UPDATE",
        list(pk.values()),
    )
    actual = dict(zip(nombres, cur.fetchone()))
    cambios = {n: v for n, v in campos.items() if actual[n] != v}
    if cambios:
        sets = ", ".join(f"{n}=%s" for n in cambios)
        cur.execute(
            f"UPDATE {tabla} SET {sets}, actualizado_en = NOW() WHERE {donde}",
            [*cambios.values(), *pk.values()],
        )
        for n, nuevo in cambios.items():
            cur.execute(
                """
                INSERT INTO ciclico_auditoria
                    (fecha, sucursal, camara, campo, valor_anterior, valor_nuevo, origen)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    pk["fecha"], pk.get("sucursal", ""), camara_log, n,
                    str(actual[n]), str(nuevo), _cfg["ip_cliente"](),
                ),
            )
        actual.update(cambios)
    return actual


def _asegurar_turno(cur, fecha):
    _upsert(
        cur, "ciclico_turnos", {"fecha": fecha},
        {
            "dia_rotacion": DIAS[fecha.weekday()],
            "ubicacion_sistema": UBICACION_DEFECTO,
            "area": AREA_DEFECTO,
        },
        {},
    )


# ----------------------------------------------------------------------------
# Lectura del estado de un día
# ----------------------------------------------------------------------------
def _leer_estado(cur, fecha):
    catalogo = _cfg["catalogo"]
    cur.execute(
        """
        SELECT dia_rotacion, desde, hasta, ubicacion_sistema, responsable, area,
               informe_novedades
        FROM ciclico_turnos WHERE fecha = %s
        """,
        (fecha,),
    )
    fila = cur.fetchone()
    if fila:
        turno = dict(zip(
            ["dia_rotacion", "desde", "hasta", "ubicacion_sistema", "responsable",
             "area", "informe_novedades"], fila))
        turno["guardado"] = True
    else:
        turno = {
            "dia_rotacion": DIAS[fecha.weekday()], "desde": "", "hasta": "",
            "ubicacion_sistema": UBICACION_DEFECTO, "responsable": "",
            "area": AREA_DEFECTO, "informe_novedades": "", "guardado": False,
        }

    cur.execute(
        "SELECT sucursal, hora_inicio, hora_fin, novedades FROM ciclico_sitios WHERE fecha = %s",
        (fecha,),
    )
    sitios_db = {r[0]: r[1:] for r in cur.fetchall()}
    cur.execute(
        """
        SELECT sucursal, orden, camara, revisada, hora_inicio, hora_fin
        FROM ciclico_camaras WHERE fecha = %s
        """,
        (fecha,),
    )
    camaras_db = {(r[0], r[1]): r[2:] for r in cur.fetchall()}

    sitios = []
    for suc, nombres in catalogo.items():
        ini, fin, nov = sitios_db.get(suc, ("", "", ""))
        camaras = []
        for i, nombre in enumerate(nombres):
            g = camaras_db.get((suc, i))
            camaras.append({
                "idx": i,
                # Si ya se guardó, se muestra el nombre tal como se revisó ese día.
                "nombre": g[0] if g else nombre,
                "revisada": bool(g[1]) if g else False,
                "hora_inicio": g[2] if g else "",
                "hora_fin": g[3] if g else "",
            })
        revisadas = sum(1 for c in camaras if c["revisada"])
        cuota = min(CUOTA_MINIMA, len(nombres))
        sitios.append({
            "sucursal": suc,
            "total": len(nombres),                 # cuadros del bloque oficial
            "total_sede": TOTAL_SEDE.get(suc, ""),  # total de cámaras de la sede (resumen)
            "cuota": cuota,
            "cuota_cumplida": cuota > 0 and revisadas >= cuota,
            "revisadas": revisadas,
            "hora_inicio": ini,
            "hora_fin": fin,
            "novedades": nov,
            "camaras": camaras,
        })
    return {"fecha": fecha.isoformat(), "turno": turno, "sitios": sitios}


def _estado_sitio(cur, fecha, sucursal):
    return next(s for s in _leer_estado(cur, fecha)["sitios"] if s["sucursal"] == sucursal)


# ----------------------------------------------------------------------------
# Validaciones
# ----------------------------------------------------------------------------
def _fecha_de(texto):
    try:
        return datetime.strptime(str(texto), "%Y-%m-%d").date()
    except ValueError:
        return None


def _hora_valida(valor, ahora_hhmm):
    """'now' -> hora actual del servidor (Santo Domingo); '' -> vacío; si no, HH:MM."""
    if valor == "now":
        return ahora_hhmm
    valor = str(valor or "").strip()
    if valor == "" or RE_HORA.match(valor):
        return valor
    raise ValueError("Hora no válida (use HH:MM)")


def _sin_db():
    if not _cfg["database_url"]:
        return jsonify(ok=False, error="Base de datos no configurada"), 503
    return None


# ----------------------------------------------------------------------------
# API
# ----------------------------------------------------------------------------
@bp.route("/api/estado")
def api_estado():
    err = _sin_db()
    if err:
        return err
    fecha = _fecha_de(request.args.get("fecha")) or _cfg["ahora"]().date()
    try:
        with _cfg["db_cursor"]() as cur:
            estado = _leer_estado(cur, fecha)
    except Exception:  # noqa: BLE001
        logger.exception("Error leyendo turno cíclico")
        return jsonify(ok=False, error="No se pudo leer"), 500
    estado["ahora"] = _cfg["ahora"]().strftime("%H:%M")
    return jsonify(ok=True, **estado)


@bp.route("/api/turno", methods=["POST"])
def api_turno():
    err = _sin_db()
    if err:
        return err
    datos = request.get_json(silent=True) or {}
    fecha = _fecha_de(datos.get("fecha"))
    if not fecha:
        return jsonify(ok=False, error="Fecha no válida"), 400
    ahora_hhmm = _cfg["ahora"]().strftime("%H:%M")
    campos = {}
    try:
        for nombre, (tipo, maximo) in CAMPOS_TURNO.items():
            if nombre not in datos:
                continue
            if tipo == "hora":
                campos[nombre] = _hora_valida(datos[nombre], ahora_hhmm)
            else:
                campos[nombre] = str(datos[nombre] or "").strip()[:maximo]
    except ValueError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    try:
        with _cfg["db_cursor"]() as cur:
            _asegurar_turno(cur, fecha)
            final = _upsert(cur, "ciclico_turnos", {"fecha": fecha}, {}, campos)
    except Exception:  # noqa: BLE001
        logger.exception("Error guardando turno cíclico")
        return jsonify(ok=False, error="No se pudo guardar"), 500
    return jsonify(ok=True, turno=final)


@bp.route("/api/sitio", methods=["POST"])
def api_sitio():
    err = _sin_db()
    if err:
        return err
    datos = request.get_json(silent=True) or {}
    fecha = _fecha_de(datos.get("fecha"))
    sucursal = datos.get("sucursal")
    if not fecha:
        return jsonify(ok=False, error="Fecha no válida"), 400
    if sucursal not in _cfg["catalogo"]:
        return jsonify(ok=False, error="Sucursal no válida"), 400
    ahora_hhmm = _cfg["ahora"]().strftime("%H:%M")
    campos = {}
    try:
        for nombre in ("hora_inicio", "hora_fin"):
            if nombre in datos:
                campos[nombre] = _hora_valida(datos[nombre], ahora_hhmm)
    except ValueError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    if "novedades" in datos:
        campos["novedades"] = str(datos["novedades"] or "").strip()[:MAX_NOVEDADES_SITIO]
    try:
        with _cfg["db_cursor"]() as cur:
            _asegurar_turno(cur, fecha)
            _upsert(cur, "ciclico_sitios", {"fecha": fecha, "sucursal": sucursal}, {}, campos)
            sitio = _estado_sitio(cur, fecha, sucursal)
    except Exception:  # noqa: BLE001
        logger.exception("Error guardando sucursal cíclica")
        return jsonify(ok=False, error="No se pudo guardar"), 500
    return jsonify(ok=True, sitio=sitio)


@bp.route("/api/camara", methods=["POST"])
def api_camara():
    err = _sin_db()
    if err:
        return err
    datos = request.get_json(silent=True) or {}
    fecha = _fecha_de(datos.get("fecha"))
    sucursal = datos.get("sucursal")
    idx = datos.get("idx")
    if not fecha:
        return jsonify(ok=False, error="Fecha no válida"), 400
    if sucursal not in _cfg["catalogo"]:
        return jsonify(ok=False, error="Sucursal no válida"), 400
    nombres = _cfg["catalogo"][sucursal]
    if not isinstance(idx, int) or isinstance(idx, bool) or not 0 <= idx < len(nombres):
        return jsonify(ok=False, error="Cámara no válida"), 400
    camara = nombres[idx]
    ahora_hhmm = _cfg["ahora"]().strftime("%H:%M")

    try:
        with _cfg["db_cursor"]() as cur:
            _asegurar_turno(cur, fecha)
            pk_sitio = {"fecha": fecha, "sucursal": sucursal}
            _upsert(cur, "ciclico_sitios", pk_sitio, {}, {})
            pk = {"fecha": fecha, "sucursal": sucursal, "orden": idx}
            _upsert(cur, "ciclico_camaras", pk, {"camara": camara}, {})
            cur.execute(
                """
                SELECT camara, revisada, hora_inicio, hora_fin FROM ciclico_camaras
                WHERE fecha=%s AND sucursal=%s AND orden=%s
                """,
                (fecha, sucursal, idx),
            )
            nombre_dia, rev_act, ini_act, fin_act = cur.fetchone()
            ini, fin, rev = ini_act, fin_act, rev_act

            if "hora_inicio" in datos:
                ini = _hora_valida(datos["hora_inicio"], ahora_hhmm)
            if "hora_fin" in datos:
                fin = _hora_valida(datos["hora_fin"], ahora_hhmm)
            if "revisada" in datos:
                rev = bool(datos["revisada"])
                if rev and not rev_act:
                    # Al marcar: la hora fin es "ahora" (salvo que se haya dado una).
                    if "hora_fin" not in datos:
                        fin = ahora_hhmm
                    # Si no tiene inicio, arranca donde terminó la cámara anterior
                    # de la misma sucursal (o al iniciar la sucursal, o ahora).
                    if not ini and "hora_inicio" not in datos:
                        cur.execute(
                            """
                            SELECT MAX(hora_fin) FROM ciclico_camaras
                            WHERE fecha=%s AND sucursal=%s AND revisada
                              AND hora_fin <> '' AND orden <> %s
                            """,
                            (fecha, sucursal, idx),
                        )
                        previa = cur.fetchone()[0]
                        cur.execute(
                            "SELECT hora_inicio FROM ciclico_sitios WHERE fecha=%s AND sucursal=%s",
                            (fecha, sucursal),
                        )
                        ini_sitio = cur.fetchone()[0]
                        ini = previa or ini_sitio or fin
                        if ini > fin:
                            ini = fin
                elif not rev and rev_act and "hora_fin" not in datos:
                    fin = ""  # al desmarcar se limpia la hora fin (queda en auditoría)

            _upsert(
                cur, "ciclico_camaras", pk, {"camara": camara},
                {"revisada": rev, "hora_inicio": ini, "hora_fin": fin},
                camara_log=nombre_dia,
            )

            # Hora de la sucursal: inicia con la primera cámara marcada y termina
            # cuando todas están marcadas (si el operador no las puso a mano).
            cur.execute(
                "SELECT COUNT(*) FILTER (WHERE revisada) FROM ciclico_camaras "
                "WHERE fecha=%s AND sucursal=%s",
                (fecha, sucursal),
            )
            marcadas = cur.fetchone()[0]
            cur.execute(
                "SELECT hora_inicio, hora_fin FROM ciclico_sitios WHERE fecha=%s AND sucursal=%s",
                (fecha, sucursal),
            )
            s_ini, s_fin = cur.fetchone()
            auto = {}
            if rev and not s_ini and ini:
                auto["hora_inicio"] = ini
            # El cuadro FIN de la sede se cierra al cumplir la cuota (ciclo 1) y se
            # corre hacia adelante si se completa el ciclo 2 (parte inferior).
            if rev and not rev_act and marcadas >= min(CUOTA_MINIMA, len(nombres)) and fin \
                    and (not s_fin or fin > s_fin):
                auto["hora_fin"] = fin
            if auto:
                _upsert(cur, "ciclico_sitios", pk_sitio, {}, auto)
            sitio = _estado_sitio(cur, fecha, sucursal)
    except ValueError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception:  # noqa: BLE001
        logger.exception("Error guardando cámara cíclica")
        return jsonify(ok=False, error="No se pudo guardar"), 500
    return jsonify(ok=True, sitio=sitio)


# ----------------------------------------------------------------------------
# Páginas
# ----------------------------------------------------------------------------
@bp.route("/")
def pagina():
    fecha = _fecha_de(request.args.get("fecha")) or _cfg["ahora"]().date()
    return render_template_string(PAGINA, fecha=fecha.isoformat())


@bp.route("/historial")
def historial():
    filas = []
    error = None
    if not _cfg["database_url"]:
        error = "La base de datos no está configurada (falta DATABASE_URL)."
    else:
        try:
            with _cfg["db_cursor"]() as cur:
                cur.execute("""
                    SELECT t.fecha, t.dia_rotacion, t.responsable, t.desde, t.hasta,
                           (SELECT COUNT(*) FROM ciclico_camaras c
                             WHERE c.fecha = t.fecha AND c.revisada),
                           (SELECT COUNT(*) FROM ciclico_sitios s
                             WHERE s.fecha = t.fecha AND s.novedades <> ''),
                           (t.informe_novedades <> '')
                    FROM ciclico_turnos t ORDER BY t.fecha DESC LIMIT 120
                """)
                filas = cur.fetchall()
        except Exception:  # noqa: BLE001
            logger.exception("Error leyendo historial cíclico")
            error = "No se pudo leer el historial."
    total = sum(len(v) for v in _cfg["catalogo"].values())
    return render_template_string(HISTORIAL, filas=filas, total=total, error=error)


@bp.route("/auditoria")
def auditoria():
    fecha = _fecha_de(request.args.get("fecha")) or _cfg["ahora"]().date()
    filas = []
    error = None
    if not _cfg["database_url"]:
        error = "La base de datos no está configurada (falta DATABASE_URL)."
    else:
        try:
            with _cfg["db_cursor"]() as cur:
                cur.execute(
                    """
                    SELECT cambiado_en, sucursal, camara, campo, valor_anterior,
                           valor_nuevo, origen
                    FROM ciclico_auditoria WHERE fecha = %s
                    ORDER BY cambiado_en DESC, id DESC LIMIT 1000
                    """,
                    (fecha,),
                )
                filas = [
                    (_cfg["local_txt"](r[0]),) + tuple(r[1:]) for r in cur.fetchall()
                ]
        except Exception:  # noqa: BLE001
            logger.exception("Error leyendo auditoría cíclica")
            error = "No se pudo leer la auditoría."
    return render_template_string(AUDITORIA, filas=filas, fecha=fecha.isoformat(), error=error)


# ----------------------------------------------------------------------------
# Descargas: PDF y Excel del turno (sólo lectura de las tablas ciclico_*)
# ----------------------------------------------------------------------------
def _txt(valor):
    """Escapa texto para Paragraph de ReportLab."""
    return escape(str(valor if valor is not None else ""))


def _fecha_larga(iso):
    return "/".join(reversed(iso.split("-")))


def _estado_para_reporte(fecha):
    """Lee el día completo y la hora del último cambio (para 'Hora del informe')."""
    with _cfg["db_cursor"]() as cur:
        estado = _leer_estado(cur, fecha)
        cur.execute("SELECT MAX(cambiado_en) FROM ciclico_auditoria WHERE fecha = %s", (fecha,))
        ultimo = cur.fetchone()[0]
    if ultimo is not None:
        # local_txt -> "dd/mm/aaaa hh:mm:ss AM"; se toma "hh:mm AM"
        partes = _cfg["local_txt"](ultimo).split(" ")
        estado["hora_informe"] = f"{partes[1][:5]} {partes[2]}"
    else:
        estado["hora_informe"] = _cfg["ahora"]().strftime("%I:%M %p")
    return estado


# --- Formato oficial INJUL (réplica del PDF institucional "Plan de Monitoreo
# Tiempo Real Cíclico Diario"). No modernizar: colores, orden de secciones,
# rótulos y estructura de bloques son normativos. ---------------------------
NAVY = colors.HexColor("#0C3B66")
VERDE = colors.HexColor("#2E7D32")
ROJO = colors.HexColor("#EA4335")
VERDE_CHECK = colors.HexColor("#08E608")
VERDE_SN = colors.HexColor("#00C800")
BORDE = colors.HexColor("#DADADA")
CAJA = colors.HexColor("#F5F5F5")
FILA_ALT = colors.HexColor("#F7F9F9")
GRIS_ETQ = colors.HexColor("#666666")
CUADRO_VACIO = colors.HexColor("#E0E0E0")

A4_W, A4_H = A4
MARGEN_X = 63
ANCHO = A4_W - 2 * MARGEN_X  # ancho útil (≈ ancho del banner oficial)

# Nombre con el que cada bloque del catálogo cíclico aparece en el formato oficial.
# El orden de este diccionario es el orden oficial de los bloques en el PDF.
NOMBRES_OFICIALES = {
    "Sucursal San Cristóbal": "SAN CRISTÓBAL",
    "Marmotech DCM": "SANTO DOMINGO (DCM)",
    "Marmotech Santiago 1": "SANTIAGO 1",
    "Marmotech Santiago 2": "SANTIAGO 2",
    "Marmotech Blue Mall": "BLUE MALL",
    "Marmotech Outlet RP": "OUTLET RP",
    "Marmotech Punta Cana": "PUNTA CANA",
    "Marmotech Caribbean": "CARIBBEAN",
    "Caribbean Agregados": "CARIBBEAN agregado",
}
ORDEN_OFICIAL = {k: i for i, k in enumerate(NOMBRES_OFICIALES)}
# Páginas del PDF oficial: 1 = datos + resumen + San Cristóbal; 2 = DCM, Santiago 1,
# Santiago 2 y Blue Mall; 3 = Outlet RP, Punta Cana, Caribbean, informe y firma;
# 4 = Caribbean agregado (se añadió después de la firma).
SALTO_ANTES_DE = {"Marmotech DCM", "Marmotech Outlet RP"}
AGREGADO_TRAS_FIRMA = "Caribbean Agregados"


def _nombre_oficial(sucursal):
    return NOMBRES_OFICIALES.get(sucursal, sucursal.upper())


def _ps(nombre, **kw):
    base = dict(fontName="Helvetica", fontSize=8, leading=10, textColor=colors.HexColor("#222222"))
    base.update(kw)
    return ParagraphStyle(nombre, **base)


E_TITULO = _ps("titulo", fontName="Helvetica-Bold", fontSize=12.5, leading=16, textColor=colors.white, alignment=TA_CENTER)
E_SUB = _ps("sub", fontSize=8.5, leading=11, textColor=colors.white, alignment=TA_CENTER)
E_ETQ = _ps("etq", fontName="Helvetica-Bold", fontSize=7, leading=9, textColor=GRIS_ETQ)
E_VAL = _ps("val", fontName="Helvetica-Bold", fontSize=8.5, leading=10)
E_BARRA = _ps("barra", fontName="Helvetica-Bold", fontSize=7.5, leading=9, textColor=colors.white)
E_TH = _ps("th", fontName="Helvetica-Bold", fontSize=6.8, leading=8, textColor=colors.white)
E_FILA = _ps("fila", fontSize=8, leading=10, textColor=colors.HexColor("#333333"))
E_SITIO = _ps("sitio", fontName="Helvetica-Bold", fontSize=10, leading=12, textColor=NAVY)
E_COL = _ps("col", fontName="Helvetica-Bold", fontSize=9, leading=11, textColor=NAVY, alignment=TA_CENTER)
E_CAM = _ps("cam", fontSize=7, leading=8.5, textColor=NAVY)
E_HORA = _ps("hora", fontName="Helvetica-Bold", fontSize=8, leading=9.5, alignment=TA_CENTER)
E_NOV_ETQ = _ps("novetq", fontSize=8, leading=10)
E_NOV = _ps("nov", fontSize=8.5, leading=11)
E_INF_T = _ps("inft", fontName="Helvetica-Bold", fontSize=8.5, leading=11, alignment=TA_CENTER)
E_INF = _ps("inf", fontName="Helvetica-Bold", fontSize=7, leading=9, alignment=TA_CENTER)
E_PIE = _ps("pie", fontName="Helvetica-Oblique", fontSize=7.5, leading=9, textColor=colors.HexColor("#888888"), alignment=TA_CENTER)


class Cuadro(Flowable):
    """Cuadro de estado (verde = ok, rojo = novedad, gris = pendiente)."""

    def __init__(self, color, lado=12):
        super().__init__()
        self.color, self.lado = color, lado

    def wrap(self, aw, ah):
        return self.lado, self.lado

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.setStrokeColor(self.color)
        c.rect(0, 0, self.lado, self.lado, stroke=0, fill=1)


class Casilla(Flowable):
    """Casilla de verificación; marcada = visto verde como en el formato oficial."""

    def __init__(self, marcada):
        super().__init__()
        self.marcada = marcada

    def wrap(self, aw, ah):
        return 12, 12

    def draw(self):
        c = self.canv
        c.setStrokeColor(colors.HexColor("#9E9E9E"))
        c.setFillColor(colors.white)
        c.setLineWidth(0.8)
        c.rect(0.5, 0.5, 10, 10, stroke=1, fill=1)
        if self.marcada:
            _visto(c, 0, 0, 1.0, 1.8)


def _visto(c, x, y, escala, grosor):
    c.saveState()
    c.setStrokeColor(VERDE_CHECK)
    c.setLineWidth(grosor)
    c.setLineCap(1)
    c.setLineJoin(1)
    p = c.beginPath()
    p.moveTo(x + 1 * escala, y + 5 * escala)
    p.lineTo(x + 4.5 * escala, y + 1 * escala)
    p.lineTo(x + 12 * escala, y + 13 * escala)
    c.drawPath(p, stroke=1, fill=0)
    c.restoreState()


def _icono_camara(c, x, y):
    """Cámara de vigilancia (icono verde del resumen)."""
    c.saveState()
    c.translate(x, y)
    c.setFillColor(VERDE)
    c.setStrokeColor(VERDE)
    c.rotate(-12)
    c.roundRect(2, 5, 13, 7.5, 2, stroke=0, fill=1)   # cuerpo
    c.roundRect(14, 6.5, 3.5, 4.5, 1, stroke=0, fill=1)  # lente
    c.setFillColor(colors.white)
    c.circle(6, 8.7, 1.5, stroke=0, fill=1)
    c.setFillColor(VERDE)
    c.setLineWidth(1.6)
    c.line(5, 5, 3, 1)                                  # soporte
    c.rect(0.5, 0, 5, 1.6, stroke=0, fill=1)
    c.restoreState()


class EstadoMonitoreo(Flowable):
    """Columna 'CÁMARAS MONITOREADAS HOY' del resumen."""

    def __init__(self, completo, revisadas, total, novedad):
        super().__init__()
        self.completo, self.revisadas, self.total, self.novedad = completo, revisadas, total, novedad

    def wrap(self, aw, ah):
        return 118, 17

    def draw(self):
        c = self.canv
        if self.completo:
            _icono_camara(c, 8, 2)
            _visto(c, 36, 1, 1.0, 1.8)
        elif self.revisadas:
            c.setFont("Helvetica", 7.5)
            c.setFillColor(GRIS_ETQ)
            c.drawString(8, 5, f"{self.revisadas} / {self.total}")
        if self.novedad:
            c.setFillColor(ROJO)
            c.rect(62, 3, 11, 11, stroke=0, fill=1)
            c.setFont("Helvetica", 6.8)
            c.drawString(77, 6, "novedad")


def _hora12(hhmm):
    if not hhmm:
        return "—"
    hh, mm = hhmm.split(":")
    n = int(hh)
    return f"{n % 12 or 12:02d}:{mm}{'AM' if n < 12 else 'PM'}"


def _turno_nombre(desde):
    if not desde:
        return ""
    h = int(desde.split(":")[0])
    return " – TURNO DE LA MAÑANA" if h < 12 else (" – TURNO DE LA TARDE" if h < 18 else " – TURNO DE LA NOCHE")


def _caja_redonda(contenido, ancho, alto=None, relleno=None, borde=BORDE, pad_izq=10):
    t = Table([[contenido]], colWidths=[ancho], rowHeights=[alto] if alto else None)
    estilo = [
        ("BOX", (0, 0), (-1, -1), 0.8, borde), ("ROUNDEDCORNERS", [3, 3, 3, 3]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), pad_izq),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    if relleno is not None:
        estilo.append(("BACKGROUND", (0, 0), (-1, -1), relleno))
    t.setStyle(TableStyle(estilo))
    return t


def _campo(etiqueta, valor, ancho):
    t = Table(
        [[Paragraph(etiqueta, E_ETQ)], [_caja_redonda(Paragraph(_txt(valor) or "&nbsp;", E_VAL), ancho, 21)]],
        colWidths=[ancho],
    )
    t.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _trozos(texto, largo=450):
    """Parte un texto largo en pedazos para que la tabla pueda saltar de página."""
    salida = []
    for linea in str(texto or "").splitlines() or [""]:
        while len(linea) > largo:
            corte = linea.rfind(" ", 0, largo)
            corte = corte if corte > largo // 2 else largo
            salida.append(linea[:corte])
            linea = linea[corte:].lstrip()
        salida.append(linea)
    return salida


def _bloque_sitio(s):
    """Bloque oficial de una sede: nombre + INICIO/FIN, cuadros de cámaras fijos con sus
    horas, y 'Novedades y acciones tomadas'.

    Las filas son SIEMPRE las del catálogo oficial (5 o 10 cuadros): la hora de cada
    cámara se estampa en su propia fila, así el ciclo 1 queda en la parte superior y el
    ciclo 2 (cámaras 6-10 de las sedes grandes) en la parte inferior.
    """
    novedad = bool(s["novedades"])
    iniciada = bool(s["hora_inicio"]) or s["revisadas"] > 0
    cerrada = bool(s["hora_fin"]) or s.get("cuota_cumplida", False)
    c_ini = ROJO if novedad else (VERDE if iniciada else CUADRO_VACIO)
    c_fin = ROJO if novedad else (VERDE if cerrada else CUADRO_VACIO)

    w0, w1 = 24, 205
    w23 = (ANCHO - w0 - w1) / 2
    base = [
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    nombre = Table([[Paragraph(_txt(_nombre_oficial(s["sucursal"])), E_SITIO), Cuadro(c_ini)]],
                   colWidths=[w1 - 40, 18])
    nombre.setStyle(TableStyle(base))
    fin_cab = Table([[Paragraph("FIN", E_COL), Cuadro(c_fin)]], colWidths=[w23 - 30, 22])
    fin_cab.setStyle(TableStyle(base + [("ALIGN", (1, 0), (1, 0), "RIGHT")]))

    filas = [[nombre, "", Paragraph("INICIO", E_COL), fin_cab]]
    for c in s["camaras"]:
        filas.append([Casilla(c["revisada"]), Paragraph(_txt(c["nombre"]), E_CAM),
                      Paragraph(_txt(c["hora_inicio"]), E_HORA), Paragraph(_txt(c["hora_fin"]), E_HORA)])
    n = len(filas) - 1  # última fila de cámaras

    filas.append([Paragraph("Novedades y acciones tomadas.", E_NOV_ETQ), "", "", ""])
    if novedad:
        cuerpo = "no:&nbsp;&nbsp;" + "<br/>".join(_txt(x) for x in _trozos(s["novedades"], 600))
    else:
        cuerpo = 'no:&nbsp;&nbsp;<font name="Helvetica-Bold" color="#00C800">s/n</font>'
    caja = _caja_redonda(Paragraph(cuerpo, E_NOV), ANCHO - 26, 28 if not novedad else None,
                         CAJA, colors.HexColor("#E3E3E3"), 12)
    filas.append([caja, "", "", ""])

    t = Table(filas, colWidths=[w0, w1, w23, w23])
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.8, BORDE), ("ROUNDEDCORNERS", [4, 4, 4, 4]),
        ("LINEBEFORE", (0, 0), (0, -1), 4, NAVY),
        ("SPAN", (0, 0), (1, 0)), ("LINEBELOW", (0, 0), (-1, 0), 0.6, BORDE),
        ("LINEBEFORE", (2, 1), (2, n), 3, colors.black), ("LINEBEFORE", (3, 1), (3, n), 3, colors.black),
        ("SPAN", (0, n + 1), (-1, n + 1)), ("SPAN", (0, n + 2), (-1, n + 2)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (2, 0), (3, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
        ("LEFTPADDING", (0, 0), (0, -1), 10),
        ("TOPPADDING", (0, 0), (-1, 0), 5), ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
        ("TOPPADDING", (0, n + 1), (-1, n + 1), 4), ("BOTTOMPADDING", (0, n + 2), (-1, n + 2), 7),
        ("LEFTPADDING", (0, n + 2), (-1, n + 2), 14), ("RIGHTPADDING", (0, n + 2), (-1, n + 2), 8),
    ]))
    return t


def _resumen(sitios_por_clave, ancho_barra):
    """'RESUMEN INFORMATIVO DEL PROCESO': las 9 sedes oficiales (Boca Chica sin datos)."""
    barra = Table([[Paragraph("RESUMEN INFORMATIVO DEL PROCESO", E_BARRA)]], colWidths=[ancho_barra])
    barra.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY), ("ROUNDEDCORNERS", [4, 4, 0, 0]),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    c1, c2 = ANCHO * 0.40, ANCHO * 0.27
    c3 = ANCHO - c1 - c2
    filas = [[Paragraph("SEDE PRINCIPAL/SUCURSALES", E_TH), Paragraph("TOTAL CÁMARAS", E_TH),
              Paragraph("CÁMARAS MONITOREADAS HOY", E_TH)]]
    for clave, rotulo, total in RESUMEN_OFICIAL:
        s = sitios_por_clave.get(clave)
        estado = (EstadoMonitoreo(s["cuota_cumplida"], s["revisadas"], s["cuota"], bool(s["novedades"]))
                  if s else "")
        filas.append([Paragraph(_txt(rotulo), E_FILA), Paragraph(_txt(total), E_FILA), estado])
    tabla = Table(filas, colWidths=[c1, c2, c3])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, FILA_ALT]),
        ("GRID", (0, 1), (-1, -1), 0.5, colors.HexColor("#E3E3E3")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (2, 1), (2, -1), "LEFT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (2, 1), (2, -1), 0),
    ]))
    return [barra, tabla]


def construir_pdf_ciclico(estado):
    turno = estado["turno"]
    sitios = sorted(estado["sitios"], key=lambda s: ORDEN_OFICIAL.get(s["sucursal"], 99))
    por_clave = {s["sucursal"]: s for s in sitios}
    fecha_txt = _fecha_larga(estado["fecha"])
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, leftMargin=MARGEN_X, rightMargin=MARGEN_X,
        topMargin=44, bottomMargin=36,
        title=f"INJUL - Plan de Monitoreo Tiempo Real Cíclico Diario {fecha_txt}", author="INJUL",
    )
    story = []

    # 1) Banner
    banner = Table([[Paragraph("CENTRO DE MONITOREO GENERAL MARMOTECH", E_TITULO)],
                    [Paragraph("INJUL · Plan de Monitoreo Tiempo Real Cíclico Diario", E_SUB)]], colWidths=[ANCHO])
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY), ("ROUNDEDCORNERS", [4, 4, 4, 4]),
        ("TOPPADDING", (0, 0), (-1, 0), 12), ("BOTTOMPADDING", (0, 0), (-1, 0), 1),
        ("TOPPADDING", (0, 1), (-1, 1), 1), ("BOTTOMPADDING", (0, 1), (-1, 1), 11),
    ]))
    story += [banner, Spacer(1, 10)]

    # 2) Datos del turno
    mitad = ANCHO / 2
    fila1 = Table([[_campo("FECHA", fecha_txt, mitad - 10), _campo("DÍA DE ROTACIÓN", turno["dia_rotacion"], mitad - 10)],
                   [_campo("DESDE", _hora12(turno["desde"]), mitad - 10), _campo("HASTA", _hora12(turno["hasta"]), mitad - 10)]],
                  colWidths=[mitad, mitad])
    fila1.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story += [fila1, _campo("UBICACIÓN / SISTEMA", turno["ubicacion_sistema"], ANCHO), Spacer(1, 3),
              _campo("RESPONSABLE DE TURNO", turno["responsable"], ANCHO), Spacer(1, 8)]

    # 3) Resumen informativo
    story += _resumen(por_clave, ANCHO - 4)
    story.append(Spacer(1, 10))

    # 4) Un bloque fijo por sede (el agregado de Caribbean va al final, tras la firma)
    for s in sitios:
        if s["sucursal"] == AGREGADO_TRAS_FIRMA:
            continue
        if s["sucursal"] in SALTO_ANTES_DE:
            story.append(PageBreak())
        story += [KeepTogether([_bloque_sitio(s)]), Spacer(1, 8)]

    # 5) Informe de novedades del turno + firma + pie
    area = turno["area"] or ""
    lineas = [f"Fecha: {fecha_txt}", f'Hora del informe: {estado.get("hora_informe", "")}',
              f"Área: {area}", "Reporte de Monitoreo"]
    cabecera = "".join(f"{_txt(x)}<br/>" for x in lineas)
    cuerpo = (turno["informe_novedades"] or "Sin observaciones generales registradas.").strip()
    nombre_resp = _txt(turno["responsable"])
    filas_inf = [[Paragraph(f"INFORME DE NOVEDADES{_turno_nombre(turno['desde'])}", E_INF_T)],
                 [Paragraph(cabecera, E_INF)]]
    # El cuerpo va en filas cortas para que un informe extenso nunca rompa el PDF.
    filas_inf += [[Paragraph(_txt(x) or "&nbsp;", E_INF)] for x in _trozos(cuerpo)]
    filas_inf.append([_caja_redonda(
        Paragraph(f"<b>{nombre_resp}</b>" if nombre_resp else "&nbsp;", _ps("r", fontName="Helvetica-Bold", fontSize=7.5)),
        ANCHO - 34, 26, CAJA, colors.HexColor("#E3E3E3"), 28)])
    n_inf = len(filas_inf)
    informe = Table(filas_inf, colWidths=[ANCHO])
    informe.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.8, BORDE), ("ROUNDEDCORNERS", [4, 4, 4, 4]),
        ("LINEBEFORE", (0, 0), (0, -1), 4, NAVY), ("LINEBELOW", (0, 0), (0, 0), 0.6, BORDE),
        ("LEFTPADDING", (0, 0), (-1, -1), 24), ("RIGHTPADDING", (0, 0), (-1, -1), 18),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, 0), 8), ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
        ("TOPPADDING", (0, 1), (-1, 1), 6), ("BOTTOMPADDING", (0, n_inf - 2), (-1, n_inf - 2), 10),
        ("TOPPADDING", (0, n_inf - 1), (-1, n_inf - 1), 4), ("BOTTOMPADDING", (0, n_inf - 1), (-1, n_inf - 1), 9),
    ]))
    firma = [
        Paragraph("FIRMA DEL RESPONSABLE", _ps("f", fontName="Helvetica-Bold", fontSize=8, textColor=GRIS_ETQ)),
        Spacer(1, 4),
        _caja_redonda(Paragraph(f"<b>{nombre_resp}</b>" if nombre_resp else "&nbsp;", _ps("fr", fontName="Helvetica-Bold", fontSize=7.5)),
                      ANCHO, 28, colors.white, BORDE, 40),
        Spacer(1, 5),
        Paragraph("Los datos se guardan en este dispositivo por fecha.", E_PIE),
    ]
    story += [informe, Spacer(1, 14)] + firma

    # 6) Bloque "Caribbean agregado": página aparte, como en el formato oficial
    if AGREGADO_TRAS_FIRMA in por_clave:
        story += [PageBreak(), _bloque_sitio(por_clave[AGREGADO_TRAS_FIRMA])]

    doc.build(story)
    buffer.seek(0)
    return buffer


def construir_excel_ciclico(estado):
    turno, sitios = estado["turno"], estado["sitios"]
    fecha_txt = _fecha_larga(estado["fecha"])
    wb = Workbook()
    cab_fill = PatternFill("solid", fgColor="0F3A63")
    cab_font = Font(bold=True, color="FFFFFF")
    ok_fill = PatternFill("solid", fgColor="EEF7EE")
    ajustar = Alignment(wrap_text=True, vertical="top")

    def cabecera(ws, fila, textos):
        for c, tx in enumerate(textos, start=1):
            celda = ws.cell(row=fila, column=c, value=tx)
            celda.fill, celda.font = cab_fill, cab_font

    # Hoja 1: Resumen del turno
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = "CENTRO DE MONITOREO GENERAL MARMOTECH — Plan de Monitoreo Cíclico Diario"
    ws["A1"].font = Font(bold=True, size=13, color="0F3A63")
    datos = [("Fecha", fecha_txt), ("Día de rotación", turno["dia_rotacion"]), ("Desde", turno["desde"]),
             ("Hasta", turno["hasta"]), ("Ubicación / Sistema", turno["ubicacion_sistema"]),
             ("Responsable de turno", turno["responsable"]), ("Área", turno["area"])]
    for i, (k, v) in enumerate(datos, start=3):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws.cell(row=i, column=2, value=v or "")
    fila = len(datos) + 4
    cabecera(ws, fila, ["Sede / Sucursal", "Cuadros del ciclo", "Revisadas", "Cuota cumplida", "Inicio", "Fin", "Novedades y acciones tomadas"])
    for s in sitios:
        fila += 1
        completo = s["cuota_cumplida"]
        valores = [s["sucursal"], s["total"], s["revisadas"], "Sí" if completo else "No",
                   s["hora_inicio"], s["hora_fin"], s["novedades"] or "s/n"]
        for c, v in enumerate(valores, start=1):
            celda = ws.cell(row=fila, column=c, value=v)
            celda.alignment = ajustar
            if completo:
                celda.fill = ok_fill
    fila += 1
    ws.cell(row=fila, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=fila, column=2, value=sum(s["total"] for s in sitios)).font = Font(bold=True)
    ws.cell(row=fila, column=3, value=sum(s["revisadas"] for s in sitios)).font = Font(bold=True)
    fila += 2
    ws.cell(row=fila, column=1, value="Informe de novedades del turno").font = Font(bold=True, color="0F3A63")
    fila += 1
    ws.cell(row=fila, column=1, value=turno["informe_novedades"] or "Sin observaciones generales registradas.").alignment = ajustar
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=7)
    ws.row_dimensions[fila].height = 90
    for col, ancho in zip("ABCDEFG", [30, 16, 12, 11, 9, 9, 60]):
        ws.column_dimensions[col].width = ancho

    # Hoja 2: Detalle por cámara
    wd = wb.create_sheet("Detalle cámaras")
    cabecera(wd, 1, ["Fecha", "Sucursal", "#", "Cámara", "Revisada", "Hora inicio", "Hora fin"])
    r = 1
    for s in sitios:
        for c in s["camaras"]:
            r += 1
            for col, v in enumerate([fecha_txt, s["sucursal"], c["idx"] + 1, c["nombre"],
                                     "Sí" if c["revisada"] else "No", c["hora_inicio"], c["hora_fin"]], start=1):
                celda = wd.cell(row=r, column=col, value=v)
                if c["revisada"]:
                    celda.fill = ok_fill
    wd.freeze_panes = "A2"
    wd.auto_filter.ref = f"A1:G{max(r, 2)}"
    for i, ancho in enumerate([12, 28, 6, 38, 10, 12, 10], start=1):
        wd.column_dimensions[get_column_letter(i)].width = ancho

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def _fecha_descarga():
    return _fecha_de(request.args.get("fecha")) or _cfg["ahora"]().date()


def _respuesta_descarga(constructor, extension, mimetype):
    if not _cfg["database_url"]:
        return Response("Base de datos no configurada (falta DATABASE_URL).", status=503)
    fecha = _fecha_descarga()
    try:
        estado = _estado_para_reporte(fecha)
        buffer = constructor(estado)
    except Exception:  # noqa: BLE001
        logger.exception("Error generando reporte cíclico (%s)", fecha.isoformat())
        return Response("No se pudo generar el reporte. Intente de nuevo.", status=500,
                        mimetype="text/plain")
    respuesta = send_file(
        buffer, mimetype=mimetype, as_attachment=True,
        download_name=f"Monitoreo_Ciclico_INJUL_{fecha.isoformat()}.{extension}",
    )
    respuesta.headers["Cache-Control"] = "no-store"  # siempre datos recién guardados
    return respuesta


@bp.route("/descargar-pdf")
def descargar_pdf():
    # PDF institucional: formato inalterable (siempre los cuadros oficiales de cada sede).
    return _respuesta_descarga(construir_pdf_ciclico, "pdf", "application/pdf")


@bp.route("/descargar-excel")
def descargar_excel():
    return _respuesta_descarga(
        construir_excel_ciclico, "xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ----------------------------------------------------------------------------
# Registro en la app principal
# ----------------------------------------------------------------------------
def registrar(app, catalogo, db_cursor, ahora, local_txt, ip_cliente, database_url):
    _cfg.update(
        catalogo=catalogo, db_cursor=db_cursor, ahora=ahora, local_txt=local_txt,
        ip_cliente=ip_cliente, database_url=database_url,
    )
    app.register_blueprint(bp)
    if database_url:
        try:
            with db_cursor() as cur:
                crear_tablas(cur)
        except Exception:  # noqa: BLE001
            logger.exception("No se pudieron crear las tablas del módulo cíclico")


ESTILO = """
<style>
  * { box-sizing: border-box; }
  body { font-family: Arial, sans-serif; background: #f4f6f9; color: #333; margin: 0; padding: 12px; }
  .wrap { max-width: 1000px; margin: 0 auto; }
  .banner { background: #0f3a63; color: #fff; border-radius: 8px; padding: 14px 16px; text-align: center; }
  .banner h1 { margin: 0; font-size: 19px; letter-spacing: .5px; }
  .banner p { margin: 4px 0 0; font-size: 13px; opacity: .9; }
  .nav { margin: 10px 0; display: flex; gap: 8px; flex-wrap: wrap; }
  .btn { display: inline-block; background: #1F497D; color: #fff; padding: 9px 14px; border: 0;
         border-radius: 4px; font-weight: bold; font-size: 13px; text-decoration: none; cursor: pointer; }
  .btn.gris { background: #6c757d; } .btn.verde { background: #2e7d32; }
  .btn.chico { padding: 5px 8px; font-size: 12px; }
  .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,.08);
          padding: 14px; margin-bottom: 14px; border-left: 4px solid #0f3a63; }
  .card h2 { margin: 0 0 10px; font-size: 15px; color: #0f3a63; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 10px; }
  label { display: block; font-size: 11px; font-weight: bold; color: #555; margin-bottom: 3px; text-transform: uppercase; }
  input[type=text], input[type=date], input[type=time], select, textarea {
    width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; background: #fff; }
  textarea { min-height: 90px; resize: vertical; font-family: inherit; }
  table { width: 100%; border-collapse: collapse; }
  th { background: #0f3a63; color: #fff; font-size: 12px; text-align: left; padding: 7px; }
  td { border-bottom: 1px solid #e3e3e3; padding: 6px 7px; font-size: 13px; }
  tr.hecho td { background: #eef7ee; }
  .ok { color: #2e7d32; font-weight: bold; } .pend { color: #b26a00; font-weight: bold; }
  .estado { position: fixed; right: 12px; bottom: 12px; background: #333; color: #fff; padding: 7px 12px;
            border-radius: 16px; font-size: 12px; opacity: .92; }
  .estado.error { background: #a94442; } .estado.ok { background: #2e7d32; }
  .descargas { display: flex; gap: 10px; flex-wrap: wrap; }
  .descargas .btn { flex: 1 1 200px; text-align: center; padding: 12px 14px; font-size: 14px; }
  .btn.rojo { background: #c0392b; }
  .btn.cargando { opacity: .6; pointer-events: none; }
  @media (max-width: 520px) { .nav .btn { flex: 1 1 100%; text-align: center; } }
  .vacio { color: #888; padding: 12px; text-align: center; }
</style>
"""

PAGINA = """<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>INJUL - Monitoreo Cíclico</title>%%ESTILO%%
<style>
  .cam { display: grid; grid-template-columns: 34px 1fr 96px 96px; gap: 6px; align-items: center;
         padding: 6px 4px; border-bottom: 1px solid #e3e3e3; }
  .cam.hecho { background: #eef7ee; }
  .cam input[type=checkbox] { width: 24px; height: 24px; cursor: pointer; }
  .cam input[type=time] { padding: 5px; font-size: 13px; }
  .cam .nom { font-size: 13px; color: #0f3a63; }
  .cam.cab { font-size: 11px; font-weight: bold; color: #555; text-transform: uppercase; }
  .barra { background: #e3e3e3; border-radius: 6px; height: 10px; overflow: hidden; margin: 6px 0 10px; }
  .barra > div { background: #2e7d32; height: 100%; width: 0; transition: width .2s; }
  .tiempos { display: flex; gap: 8px; align-items: flex-end; flex-wrap: wrap; margin-bottom: 8px; }
  .tiempos > div { width: 130px; }
  @media (max-width: 520px) { .cam { grid-template-columns: 34px 1fr 82px 82px; } }
</style></head><body><div class="wrap">
  <div class="banner">
    <h1>CENTRO DE MONITOREO GENERAL MARMOTECH</h1>
    <p>INJUL · Plan de Monitoreo Tiempo Real Cíclico Diario</p>
  </div>
  <div class="nav">
    <a class="btn gris" href="/">← Consolidado CCTV</a>
    <a class="btn gris" href="/ciclico/historial">Historial de turnos</a>
    <a class="btn gris" id="lnk-aud" href="#">Auditoría del día</a>
    <a class="btn rojo" id="dl-pdf-top" href="#">⬇ PDF oficial</a>
    <a class="btn verde" id="dl-xlsx-top" href="#">⬇ Excel</a>
  </div>

  <div class="card"><h2>Datos del turno</h2>
    <div class="grid">
      <div><label>Fecha</label><input type="date" id="f-fecha" value="{{ fecha }}"></div>
      <div><label>Día de rotación</label><input type="text" id="t-dia_rotacion" maxlength="40"></div>
      <div><label>Desde</label><input type="time" id="t-desde"></div>
      <div><label>Hasta</label><input type="time" id="t-hasta"></div>
      <div><label>Ubicación / Sistema</label><input type="text" id="t-ubicacion_sistema" maxlength="150"></div>
      <div><label>Responsable de turno</label><input type="text" id="t-responsable" maxlength="120"></div>
    </div>
  </div>

  <div class="card"><h2>Resumen informativo del proceso</h2>
    <table><thead><tr><th>Sede / Sucursal</th><th>Cuadros del ciclo</th><th>Revisadas hoy</th>
      <th>Inicio</th><th>Fin</th><th>Novedad</th></tr></thead>
      <tbody id="resumen"></tbody></table>
  </div>

  <div class="card"><h2>Revisión por sucursal</h2>
    <label>Sucursal</label><select id="sel-suc"></select>
    <div id="panel" style="margin-top:12px"></div>
  </div>

  <div class="card"><h2>Informe de novedades del turno</h2>
    <div class="grid" style="margin-bottom:10px">
      <div><label>Área</label><input type="text" id="t-area" maxlength="150"></div>
      <div><label>Fecha del informe</label><input type="text" id="i-fecha" disabled></div>
    </div>
    <label>Observaciones generales de operatividad (retrasos, lags, fallas de transmisión…)</label>
    <textarea id="t-informe_novedades" style="min-height:160px"></textarea>
    <div style="margin-top:10px"><label>Firma del responsable</label>
      <input type="text" id="i-firma" disabled></div>
    <p style="font-size:12px;color:#777;margin:8px 0 0">La firma toma el nombre del responsable de turno.</p>
  </div>
  <div class="card"><h2>Descargar reporte del turno</h2>
    <p style="font-size:12px;color:#777;margin:0 0 10px">El PDF sigue el formato oficial INJUL: datos del turno, resumen, bloque por sucursal con horas de inicio y fin, novedades, informe de turno y firma, de la fecha seleccionada.</p>
    <div class="descargas">
      <a class="btn rojo" id="dl-pdf" href="#">⬇ Descargar PDF oficial</a>
      <a class="btn verde" id="dl-xlsx" href="#">⬇ Descargar Excel</a>
    </div>
  </div>
  <div id="estado" class="estado" style="display:none"></div>
</div>
<script>
var S = {fecha: "{{ fecha }}", d: null, sel: 0, pend: 0, ahora: ""};
var CAMPOS = ["dia_rotacion","desde","hasta","ubicacion_sistema","responsable","area","informe_novedades"];
function $(id){ return document.getElementById(id); }
function el(tag, props, kids){
  var e = document.createElement(tag);
  for (var k in (props||{})) { if (k === "text") e.textContent = props[k]; else if (k === "cls") e.className = props[k]; else e[k] = props[k]; }
  (kids||[]).forEach(function(c){ e.appendChild(c); });
  return e;
}
function aviso(txt, clase){
  var b = $("estado"); b.style.display = "block"; b.className = "estado " + (clase||""); b.textContent = txt;
  if (clase === "ok") { setTimeout(function(){ if (S.pend === 0 && b.className.indexOf("ok") > -1) b.style.display = "none"; }, 1500); }
}
function post(url, body){
  S.pend++; aviso("Guardando…");
  return fetch(url, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(body)})
    .then(function(r){ return r.json().then(function(j){ return {ok:r.ok, j:j}; }); })
    .then(function(res){
      S.pend--;
      if (!res.ok || !res.j.ok) throw new Error(res.j.error || "Error");
      aviso("Guardado ✓", "ok"); return res.j;
    }).catch(function(e){ S.pend--; aviso("NO guardado: " + e.message + " — reintenta", "error"); throw e; });
}
function debounce(fn, ms){ var t; var f = function(){ clearTimeout(t); t = setTimeout(fn, ms); }; f.flush = function(){ clearTimeout(t); fn(); }; return f; }

function cargar(){
  fetch("/ciclico/api/estado?fecha=" + S.fecha).then(function(r){ return r.json(); }).then(function(j){
    if (!j.ok) { aviso(j.error || "Error al cargar", "error"); return; }
    S.d = j; S.ahora = j.ahora;
    CAMPOS.forEach(function(c){ $("t-" + c).value = j.turno[c] || ""; });
    $("i-fecha").value = S.fecha.split("-").reverse().join("/");
    $("i-firma").value = j.turno.responsable || "";
    $("lnk-aud").href = "/ciclico/auditoria?fecha=" + S.fecha;
    var sel = $("sel-suc"); sel.innerHTML = "";
    j.sitios.forEach(function(s, i){ sel.appendChild(el("option", {value: i, text: s.sucursal})); });
    sel.value = Math.min(S.sel, j.sitios.length - 1);
    pintarResumen(); pintarPanel();
  }).catch(function(){ aviso("No se pudo cargar el turno", "error"); });
}

function pintarResumen(){
  var tb = $("resumen"); tb.innerHTML = "";
  S.d.sitios.forEach(function(s, i){
    var completo = s.cuota_cumplida;
    var tr = el("tr", {cls: completo ? "hecho" : ""}, [
      el("td", {text: s.sucursal + (s.total_sede ? " (" + s.total_sede + " cám.)" : "")}), el("td", {text: s.total}),
      el("td", {cls: completo ? "ok" : (s.revisadas ? "pend" : ""), text: s.revisadas + " / " + s.total + (completo ? " ✓" : "")}),
      el("td", {text: s.hora_inicio || "—"}), el("td", {text: s.hora_fin || "—"}),
      el("td", {text: s.novedades ? "Sí" : "s/n"})]);
    tr.style.cursor = "pointer";
    tr.onclick = function(){ S.sel = i; $("sel-suc").value = i; pintarPanel(); };
    tb.appendChild(tr);
  });
}

function reemplazarSitio(sitio){
  for (var i = 0; i < S.d.sitios.length; i++) if (S.d.sitios[i].sucursal === sitio.sucursal) {
    var nov = S.d.sitios[i].novedades;   // las novedades en edición no se pisan
    S.d.sitios[i] = sitio; sitio.novedades = nov;
  }
  pintarResumen();
  if (S.d.sitios[S.sel].sucursal === sitio.sucursal) pintarChecklist();
}

function guardarCamara(s, c, extra){
  var body = {fecha: S.fecha, sucursal: s.sucursal, idx: c.idx};
  for (var k in extra) body[k] = extra[k];
  return post("/ciclico/api/camara", body).then(function(j){ reemplazarSitio(j.sitio); })
    .catch(function(){ pintarChecklist(); });
}
function sitioVigente(nombre){
  for (var i = 0; i < S.d.sitios.length; i++) if (S.d.sitios[i].sucursal === nombre) return S.d.sitios[i];
}
function guardarSitio(s, extra){
  var body = {fecha: S.fecha, sucursal: s.sucursal};
  for (var k in extra) body[k] = extra[k];
  return post("/ciclico/api/sitio", body).then(function(j){
    if (!("novedades" in extra)) reemplazarSitio(j.sitio);
    else pintarResumen();
  });
}

function pintarPanel(){
  var s = S.d.sitios[S.sel], p = $("panel"); p.innerHTML = "";
  var ini = el("input", {type:"time", value: s.hora_inicio || "", id:"s-ini"});
  var fin = el("input", {type:"time", value: s.hora_fin || "", id:"s-fin"});
  ini.onchange = function(){ guardarSitio(s, {hora_inicio: ini.value}); };
  fin.onchange = function(){ guardarSitio(s, {hora_fin: fin.value}); };
  var bIni = el("button", {cls:"btn chico verde", text:"▶ Iniciar ahora", type:"button"});
  bIni.onclick = function(){ guardarSitio(s, {hora_inicio: "now"}); };
  var bFin = el("button", {cls:"btn chico", text:"■ Finalizar ahora", type:"button"});
  bFin.onclick = function(){ guardarSitio(s, {hora_fin: "now"}); };
  p.appendChild(el("div", {cls:"tiempos"}, [
    el("div", {}, [el("label", {text:"Hora de inicio"}), ini]), bIni,
    el("div", {}, [el("label", {text:"Hora de fin"}), fin]), bFin]));
  p.appendChild(el("div", {id:"prog-txt", style:"font-size:13px;font-weight:bold;color:#0f3a63"}));
  p.appendChild(el("div", {cls:"barra"}, [el("div", {id:"prog-bar"})]));
  var ver = el("label", {style:"text-transform:none;font-size:13px;font-weight:normal;margin-bottom:8px"});
  var chk = el("input", {type:"checkbox", id:"solo-pend"}); chk.onchange = pintarChecklist;
  ver.appendChild(chk); ver.appendChild(document.createTextNode(" Mostrar sólo pendientes"));
  p.appendChild(ver);
  p.appendChild(el("div", {id:"lista"}));
  p.appendChild(el("label", {text:"Novedades y acciones tomadas — " + s.sucursal, style:"margin-top:12px"}));
  var ta = el("textarea", {id:"s-nov", placeholder:"Escriba aquí cualquier anomalía (vacío = sin novedad)", value: s.novedades || ""});
  var guardarNov = debounce(function(){ guardarSitio(s, {novedades: ta.value}); }, 800);
  ta.oninput = function(){ sitioVigente(s.sucursal).novedades = ta.value.trim(); guardarNov(); }; ta.onblur = function(){ guardarNov.flush(); };
  p.appendChild(ta);
  pintarChecklist();
}

function pintarChecklist(){
  var s = S.d.sitios[S.sel], lista = $("lista"); if (!lista) return;
  $("s-ini").value = s.hora_inicio || ""; $("s-fin").value = s.hora_fin || "";
  $("prog-txt").textContent = "Cámaras revisadas: " + s.revisadas + " de " + s.total;
  $("prog-bar").style.width = (s.total ? 100 * s.revisadas / s.total : 0) + "%";
  var soloPend = $("solo-pend").checked;
  lista.innerHTML = "";
  lista.appendChild(el("div", {cls:"cam cab"}, [el("span"), el("span", {text:"Cámara (EZStation)"}), el("span", {text:"Inicio"}), el("span", {text:"Fin"})]));
  var nCiclos = s.total > 5 ? 2 : 1;
  s.camaras.forEach(function(c){
    if (nCiclos === 2 && (c.idx === 0 || c.idx === 5)) {
      var arriba = c.idx === 0;
      var ini = arriba ? 0 : 5, hecho = 0;
      for (var k = ini; k < ini + 5 && k < s.camaras.length; k++) if (s.camaras[k].revisada) hecho++;
      lista.appendChild(el("div", {cls:"cam cab", style:"background:#eef2f7;border-radius:4px;grid-template-columns:1fr"},
        [el("span", {text: (arriba ? "Ciclo 1 · parte superior (cámaras 1–5)" : "Ciclo 2 · parte inferior (cámaras 6–10)") + " — " + hecho + " / 5"})]));
    }
    if (soloPend && c.revisada) return;
    var cb = el("input", {type:"checkbox", checked: c.revisada});
    cb.onchange = function(){ guardarCamara(s, c, {revisada: cb.checked}); };
    var ci = el("input", {type:"time", value: c.hora_inicio || ""});
    ci.onchange = function(){ guardarCamara(s, c, {hora_inicio: ci.value}); };
    var cf = el("input", {type:"time", value: c.hora_fin || ""});
    cf.onchange = function(){ guardarCamara(s, c, {hora_fin: cf.value}); };
    lista.appendChild(el("div", {cls:"cam" + (c.revisada ? " hecho" : "")}, [cb, el("span", {cls:"nom", text:c.nombre}), ci, cf]));
  });
}

function enlazarDescargas(){
  [["dl-pdf","descargar-pdf"],["dl-xlsx","descargar-excel"],["dl-pdf-top","descargar-pdf"],["dl-xlsx-top","descargar-excel"]].forEach(function(par){
    var a = $(par[0]), base = "/ciclico/" + par[1] + "?fecha=" + S.fecha;
    var armar = function(){ return base; };
    a.href = armar();
    a.onclick = function(ev){
      ev.preventDefault();
      var url = armar();
      if (a.classList.contains("cargando")) return;
      a.classList.add("cargando");
      var texto = a.textContent; a.textContent = "Preparando…";
      var t0 = Date.now();
      // Si hay cambios sin terminar de guardar, se espera (máx. 8 s) para que el reporte salga completo.
      var espera = setInterval(function(){
        if (S.pend === 0 || Date.now() - t0 > 8000) {
          clearInterval(espera);
          window.location.href = url;
          setTimeout(function(){ a.classList.remove("cargando"); a.textContent = texto; }, 2500);
        }
      }, 150);
    };
  });
}
function enlazarTurno(){
  CAMPOS.forEach(function(c){
    var inp = $("t-" + c);
    var enviar = function(){
      var b = {fecha: S.fecha}; b[c] = inp.value;
      post("/ciclico/api/turno", b).then(function(j){
        if (c === "responsable") $("i-firma").value = inp.value;
      });
    };
    if (inp.tagName === "TEXTAREA") {
      var d = debounce(enviar, 800); inp.oninput = d; inp.onblur = function(){ d.flush(); };
    } else { inp.onchange = enviar; }
  });
  $("f-fecha").onchange = function(){
    if (!this.value) return;
    window.location.href = "/ciclico/?fecha=" + this.value;
  };
  $("sel-suc").onchange = function(){ S.sel = parseInt(this.value, 10); pintarPanel(); };
  window.addEventListener("beforeunload", function(e){ if (S.pend > 0) { e.preventDefault(); e.returnValue = ""; } });
}
enlazarTurno(); enlazarDescargas(); cargar();
</script></body></html>
"""

HISTORIAL = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>INJUL - Historial cíclico</title>%%ESTILO%%</head><body><div class="wrap">
  <div class="banner"><h1>HISTORIAL DE TURNOS CÍCLICOS</h1><p>INJUL · Plan de Monitoreo Tiempo Real Cíclico Diario</p></div>
  <div class="nav"><a class="btn" href="/ciclico/">← Turno de hoy</a><a class="btn gris" href="/">Consolidado CCTV</a></div>
  <div class="card">
  {% if error %}<div class="vacio">{{ error }}</div>
  {% elif not filas %}<div class="vacio">Aún no hay turnos registrados.</div>
  {% else %}
  <div style="overflow-x:auto"><table>
    <tr><th>Fecha</th><th>Día</th><th>Responsable</th><th>Horario</th><th>Cámaras revisadas</th>
        <th>Sucursales con novedad</th><th>Informe</th><th></th></tr>
    {% for f in filas %}<tr>
      <td>{{ f[0].strftime('%d/%m/%Y') }}</td><td>{{ f[1] }}</td><td>{{ f[2] or '—' }}</td>
      <td>{{ f[3] or '—' }} – {{ f[4] or '—' }}</td>
      <td class="{{ 'ok' if f[5] >= total else 'pend' }}">{{ f[5] }} / {{ total }}</td>
      <td>{{ f[6] }}</td><td>{{ 'Sí' if f[7] else '—' }}</td>
      <td><a href="/ciclico/?fecha={{ f[0].isoformat() }}">Abrir</a> ·
          <a href="/ciclico/auditoria?fecha={{ f[0].isoformat() }}">Auditoría</a> ·
          <a href="/ciclico/descargar-pdf?fecha={{ f[0].isoformat() }}">PDF</a> ·
          <a href="/ciclico/descargar-excel?fecha={{ f[0].isoformat() }}">Excel</a></td>
    </tr>{% endfor %}
  </table></div>{% endif %}
  </div></div></body></html>
"""

AUDITORIA = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>INJUL - Auditoría cíclica</title>%%ESTILO%%</head><body><div class="wrap">
  <div class="banner"><h1>AUDITORÍA DEL TURNO {{ fecha }}</h1><p>Cada cambio registrado, con hora y origen</p></div>
  <div class="nav"><a class="btn" href="/ciclico/?fecha={{ fecha }}">← Volver al turno</a>
    <a class="btn gris" href="/ciclico/historial">Historial</a></div>
  <div class="card">
  {% if error %}<div class="vacio">{{ error }}</div>
  {% elif not filas %}<div class="vacio">Sin cambios registrados para esta fecha.</div>
  {% else %}
  <div style="overflow-x:auto"><table>
    <tr><th>Cuándo</th><th>Sucursal</th><th>Cámara</th><th>Campo</th><th>Antes</th><th>Después</th><th>Origen</th></tr>
    {% for f in filas %}<tr><td>{{ f[0] }}</td><td>{{ f[1] or '—' }}</td><td>{{ f[2] or '—' }}</td>
      <td>{{ f[3] }}</td><td>{{ f[4] }}</td><td>{{ f[5] }}</td><td>{{ f[6] }}</td></tr>{% endfor %}
  </table></div>{% endif %}
  </div></div></body></html>
"""


PAGINA = PAGINA.replace("%%ESTILO%%", ESTILO)
HISTORIAL = HISTORIAL.replace("%%ESTILO%%", ESTILO)
AUDITORIA = AUDITORIA.replace("%%ESTILO%%", ESTILO)
