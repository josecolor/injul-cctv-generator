"""INJUL - Plan de Monitoreo Cíclico y Novedades (módulo independiente).

Es un Blueprint de Flask con sus propias tablas (prefijo ``ciclico_``), sus propias
rutas (``/ciclico``) y su propia interfaz. No toca los reportes consolidados, el
estado de cámaras ni los tickets del sistema principal: sólo LEE el catálogo
oficial de cámaras (``CATALOGO_CAMARAS``) que le entrega ``main.py``.

Todo se guarda en PostgreSQL al instante, vinculado por fecha, y cada cambio queda
en ``ciclico_auditoria`` para no perder el rastro de nada.
"""

import logging
import re
from datetime import date, datetime

from flask import Blueprint, jsonify, render_template_string, request

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
        sitios.append({
            "sucursal": suc,
            "total": len(nombres),
            "revisadas": sum(1 for c in camaras if c["revisada"]),
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
            if marcadas == len(nombres) and not s_fin and fin:
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
    <table><thead><tr><th>Sede / Sucursal</th><th>Total cámaras</th><th>Revisadas hoy</th>
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
    var completo = s.revisadas === s.total;
    var tr = el("tr", {cls: completo ? "hecho" : ""}, [
      el("td", {text: s.sucursal}), el("td", {text: s.total}),
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
  s.camaras.forEach(function(c){
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
enlazarTurno(); cargar();
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
          <a href="/ciclico/auditoria?fecha={{ f[0].isoformat() }}">Auditoría</a></td>
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
