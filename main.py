"""INJUL - Consolidado Maestro de CCTV.
Flask + PostgreSQL + Redis (opcional).
"""

import json
import logging
import os
import re
from contextlib import contextmanager
from datetime import date, datetime
from io import BytesIO
from xml.sax.saxutils import escape

import openpyxl
import psycopg2
import redis
from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.graphics.charts.barcharts import HorizontalBarChart
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

try:
    from zoneinfo import ZoneInfo
    ZONA_HORARIA = ZoneInfo("America/Santo_Domingo")
except Exception:
    from datetime import timedelta, timezone
    ZONA_HORARIA = timezone(timedelta(hours=-4), "AST")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("injul-cctv")

app = Flask(__name__)

DATABASE_URL = os.environ.get("DATABASE_URL")
REDIS_URL = os.environ.get("REDIS_URL")
CACHE_KEY = "cctv:datos_reporte"
CACHE_TTL = 300

cache = None
if REDIS_URL:
    try:
        cache = redis.Redis.from_url(
            REDIS_URL, socket_connect_timeout=2, socket_timeout=2, decode_responses=True
        )
    except Exception:
        cache = None

CATALOGO_CAMARAS = {
    "Caribbean Agregados": [
        "Frente Oficina", "Puerta desde CR", "Lado Atras", "IP Camera 04",
        "Frente Lado desde Caribbean", "Lado Atras esquina puerta", "Puerta lado norte",
        "Frente hacia lado norte", "Oficina Gabinete data", "IP Camera 10", "IP Camera 11",
        "IP Camera 12", "IP Camera 13", "IP Camera 14", "IP Camera 15", "IP Camera 16",
    ],
    "Marmotech Blue Mall": [
        "Show 2da planta", "Area Clientes 1ra planta", "Area 1ra planta puerta trasera",
        "2da Planta banos", "Area 1ra planta puerta trasera (2)", "Gabinete TI",
        "IP Camera 07", "2da Planta banos (2)", "IP Camera 09", "IP Camera 10",
        "IP Camera 11", "IP Camera 12", "IP Camera 13", "IP Camera 14", "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech DCM": [
        "Puerta Principal", "Cocina", "IP Camera 03", "S1 Entrada Vehiculos",
        "S2 Pasillo lateral", "IP Camera 06", "1ra planta asensor visita", "IP Camera 09",
        "IP Camera 11", "segundo nivel banera", "seg piso griferia", "IP Camera 14",
        "IP Camera 15", "Entrada edificio", "S2 Salida", "IP Camera 18", "IP Camera 19",
        "5to piso onix miel", "S2 Salida V", "S2 Salida Planta E", "Entrada Calle Parqueos",
        "Frente Edificio", "Salida S1 Vehiculos", "S1 Entrada al S2",
        "S1 Entrada vista de frente", "S2 Area Piezas y Vehiculos", "IP Camera 30",
        "almacen griferia piiso 1", "piso 1 entrada plancha", "IP Camera 33",
        "IP Camera 34", "IP Camera 35", "IP Camera 36", "IP Camera 37", "IP Camera 38",
        "IP Camera 39", "IP Camera 40", "IP Camera 41", "IP Camera 42", "IP Camera 43",
        "IP Camera 45", "Oficina Almacen", "Entrada a cobro lobby", "IP Camera 49",
        "Salida Asensor c", "IP Camera 51", "IP Camera 52", "IP Camera 53", "IP Camera 54",
        "IP Camera 55", "IP Camera 56",
    ],
    "Marmotech Outlet RP": [
        "IP Camera 01", "IP Camera 02", "IP Camera 03", "IP Camera 04", "IP Camera 05",
        "IP Camera 06", "IP Camera 07", "IP Camera 08", "IP Camera 09", "IP Camera 10",
        "IP Camera 11", "IP Camera 12", "IP Camera 13", "IP Camera 14", "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech Punta Cana": [
        "Puerta Entrada Principal", "Area trasera 1", "Area Trasera 2", "Area Trasera 5",
        "Area trasera 3", "Area Trasera 4", "Oficina 1", "Area Gabinete datos", "Almacen",
        "Area Lateral Derecha", "Zona Carga / Descarga", "Almacen 1", "Frente Edificacion",
        "Instalacion", "Entrada Oficinas", "Almacen 2 Escalera Subir Techo Oficina",
        "Almacen 3",
    ],
    "Sucursal San Cristóbal": [
        "Almacen Repuesto", "Parqueo 1 Tienda", "Entrada Tienda S", "Parqueo 2 Tienda",
        "Almacen 1 Tienda", "C/Entrada Tienda", "Almacen 2 Tienda", "Almacen 3 Tienda",
        "Calle Entrada 2", "Oficina Tienda S", "Parqueo 3 Tiend", "Almacen 4 Tienda",
        "Almacen 5 Tienda", "Area Caja Oficin", "MODULARES 1",
        "Entrada Mantenimiento norte", "MODULARES Area F", "MODULARES 4", "Alm. Gene 04",
        "Alm Gene 06", "Cuarto de Compresores", "Alm Gene 01", "Alm Gene 02", "MOTORES 2",
        "ENTRADA P1", "Puerta Entrada", "ARTECH 02", "PRODUCCION 1", "ARTECH 01",
        "Alm Gene 03", "PULIDORAS", "ARTECH03", "PULIDORA 02", "DATA CENTER 1",
        "Carga Area 1 Alm", "Puerta Entrada A", "Area Motores Alm", "Carga 2 Alm Gene",
        "Gasoil entrada /", "Area Parqueo 1 F", "Salida Fabrica A", "Gasoil Dispensad",
        "Gasoil", "Parqueo Fabrica", "Puerta Entrada G", "Garita Guardian",
        "Escritorio Garit", "Frente Puerta no", "Almacen Repuesto", "Almacen Repuesto (2)",
        "Ing y Calidad", "Alm Gene 07", "Area de Planta", "ARTECH 1", "ARTECH 2",
        "Alm Gene 05", "IP Camera 57", "IP Camera 58", "IP Camera 59",
    ],
    "Marmotech Santiago 1": [
        "Almacen Repuesto", "Parqueo 1 Tienda", "Entrada Tienda 5", "Parqueo 2 Tienda",
        "Almacen 1 Tienda", "C/Entrada Tienda", "Almacen 2 Tienda", "Almacen 3 Tienda",
        "Calle Entrada 2", "Oficina Tienda S", "Parqueo 3 Tiend", "Almacen 4 Tienda",
        "Almacen 5 Tienda", "Area Caja Oficin", "MODULARES 1",
        "Entrada Mantenimiento norte",
    ],
    "Marmotech Santiago 2": [
        "IP Camera 01", "IP Camera 02", "IP Camera 03", "IP Camera 04", "IP Camera 05",
        "IP Camera 06", "IP Camera 07", "IP Camera 08", "IP Camera 09", "IP Camera 10",
        "IP Camera 11", "IP Camera 12", "IP Camera 13", "IP Camera 14", "IP Camera 15",
        "IP Camera 16",
    ],
    "Santiago 2 - Sección Muebles": [
        "MUEBLES", "ENTRADA", "COCINA", "GRIFERIA", "PARQUEO O-E", "BANOS", "PISOS",
        "SALA REUNION", "PARQUEO E-O", "COUNTER", "MASETEROS", "IP Camera 12",
        "IP Camera 13", "IP Camera 14", "IP Camera 15", "IP Camera 16",
    ],
}

CATALOGO_CICLICO = {
    "Sucursal San Cristóbal": [
        "Entrada puerta 1 - entrada p1", "Área de diseño artech - artech02",
        "Área de diseño artech 2 - artech01", "Pulidoras", "Pulidora 02",
        "Carga almacén general - carga área 1 alm", "Carga almacén general 2",
        "Entrada a fábrica puerta 3 - salida fábrica A",
        "Garita punchado - puerta entrada G", "Calle principal puerta 2 - frente puerta",
    ],
    "Marmotech DCM": [
        "IP Camera 30", "Frente de la puerta, utensilios - IP Camera 32",
        "Parqueo de motores - IP Camera 34", "Terraza lado céntrico - IP Camera 36",
        "IP Camera 38", "Parqueo sótano 1 - IP Camera 40",
        "Pasillos de exhibidores parte céntrica - IP Camera 42", "IP Camera 44",
        "IP Camera 46", "IP Camera 48",
    ],
    "Marmotech Santiago 1": [
        "Calle frente a la empresa - IP Camera 02", "Parte trasera - IP Camera 04",
        "Calle principal - IP Camera 06", "Mostrador - IP Camera 08",
        "Almacén interior 2 - IP Camera 10",
    ],
    "Marmotech Santiago 2": [
        "Mostrador - Muebles", "Cocina", "Parqueo principal - Parqueo 0-E", "Pisos",
        "Parqueo E-O",
    ],
    "Marmotech Blue Mall": [
        "Show 2da Planta", "Área 1ra Planta Puerta Trasera",
        "Área 1ra Planta Puerta Trasera Norte", "Área Clientes 1ra Planta", "Gabinete TI",
    ],
    "Marmotech Outlet RP": [
        "Parqueo lado derecho - IP Camera 10", "Parte atrás, compresores A/C - IP Camera 12",
        "Almacén General - IP Camera 14", "Exhibidores de venta - IP Camera 16",
        "Cámara escalera lado izquierdo - IP Camera 02",
    ],
    "Marmotech Punta Cana": [
        "Área Trasera 1", "Área Trasera 5", "Área Trasera 4", "Área Gabinete Datos",
        "Puerta Entrada Principal",
    ],
    "Marmotech Caribbean": [
        "CB 6-7-8 Planta Eléctrica", "Taller Soldar", "Oficina Almacén Repuesto",
        "Almacén Repuesto", "Taller Fabricación",
    ],
    "Caribbean Agregados": [
        "Frente Lado desde Caribbean", "Lado Atrás esquina puerta", "Puerta lado norte",
        "Frente hacia lado norte", "Oficina Gabinete data",
    ],
}

ANCLA_ROTACION = date(2026, 10, 24)
INVENTARIO_DE = {k: k for k in CATALOGO_CAMARAS}

def _construir_pool(clave):
    base = list(CATALOGO_CICLICO[clave])
    vistos = {n.lower() for n in base}
    nums = {m for n in base for m in re.findall(r"ip camera\s*0*(\d+)", n.lower())}
    extra = []
    for n in CATALOGO_CAMARAS.get(INVENTARIO_DE.get(clave, ""), []):
        ip = re.findall(r"ip camera\s*0*(\d+)$", n.lower())
        if n.lower() in vistos or (ip and ip[0] in nums):
            continue
        vistos.add(n.lower())
        extra.append(n)
    return base + extra

POOL_CICLICO = {clave: _construir_pool(clave) for clave in CATALOGO_CICLICO}

def camaras_no_visibles(cur, sucursal):
    cur.execute(
        "SELECT camara FROM estado_camaras_actual WHERE sucursal = %s AND estado = 'No visible'",
        (sucursal,),
    )
    return {fila[0] for fila in cur.fetchall()}

ESTADOS = [
    ("Operativa", "Operativa (Sí)"),
    ("No visible", "No visible (Rojo)"),
    ("Por verificar", "Por verificar"),
]
ESTADOS_VALIDOS = {valor for valor, _ in ESTADOS}
CLASES_ESTADO = {"Operativa": "estado-op", "No visible": "estado-nv", "Por verificar": "estado-pv"}
MENSAJES = {
    "ticket_ok": ("ok", "Ticket registrado correctamente."),
    "sin_db": ("error", "Base de datos no configurada."),
    "error": ("error", "No se pudo guardar."),
}

def ahora():
    return datetime.now(ZONA_HORARIA) if ZONA_HORARIA else datetime.now()

def local_txt(momento):
    if momento is None:
        return ""
    if ZONA_HORARIA and momento.tzinfo:
        momento = momento.astimezone(ZONA_HORARIA)
    return momento.strftime("%d/%m/%Y %I:%M:%S %p")

def esc(texto):
    return escape(str(texto if texto is not None else ""))

def ip_cliente():
    reenviado = request.headers.get("X-Forwarded-For", "")
    return (reenviado.split(",")[0].strip() or request.remote_addr or "")[:60]

@contextmanager
def db_cursor():
    conn = psycopg2.connect(DATABASE_URL, connect_timeout=5)
    try:
        with conn:
            with conn.cursor() as cur:
                yield cur
    finally:
        conn.close()

def inicializar_base_de_datos():
    if not DATABASE_URL:
        return
    try:
        with db_cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS reportes_tickets (
                    id SERIAL PRIMARY KEY,
                    fecha DATE NOT NULL,
                    ticket VARCHAR(50) NOT NULL,
                    sucursal_area VARCHAR(150) NOT NULL,
                    estado VARCHAR(50) NOT NULL,
                    dias_abiertos INT DEFAULT 0,
                    detalle TEXT,
                    creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS estado_camaras_actual (
                    sucursal VARCHAR(100) NOT NULL,
                    camara VARCHAR(150) NOT NULL,
                    estado VARCHAR(50) NOT NULL,
                    en_vivo VARCHAR(20) NOT NULL DEFAULT 'Sí',
                    grabacion VARCHAR(30) NOT NULL DEFAULT 'Normal',
                    observacion TEXT NOT NULL DEFAULT '',
                    actualizado_en TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    PRIMARY KEY (sucursal, camara)
                );
                CREATE TABLE IF NOT EXISTS historial_cctv (
                    id BIGSERIAL PRIMARY KEY,
                    cambiado_en TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    sucursal VARCHAR(100) NOT NULL,
                    camara VARCHAR(150) NOT NULL,
                    estado_anterior VARCHAR(50),
                    estado_nuevo VARCHAR(50) NOT NULL,
                    observacion_anterior TEXT,
                    observacion_nueva TEXT,
                    origen VARCHAR(60)
                );
            """)
    except Exception:
        logger.exception("Error inicializando BD")

def cargar_datos():
    datos = {"detalles": {}, "tickets": [], "actualizado": ""}
    if not DATABASE_URL:
        return datos
    try:
        with db_cursor() as cur:
            cur.execute("SELECT sucursal, camara, estado, en_vivo, grabacion, observacion, actualizado_en FROM estado_camaras_actual")
            ultima = None
            for suc, cam, est, env, grb, obs, momento in cur.fetchall():
                datos["detalles"].setdefault(suc, {})[cam] = {
                    "estado": est,
                    "en_vivo": env or "",
                    "grabacion": grb or "Normal",
                    "observacion": obs or "",
                }
                if ultima is None or momento > ultima:
                    ultima = momento
            datos["actualizado"] = local_txt(ultima)
            cur.execute("SELECT ticket, fecha, sucursal_area, estado, dias_abiertos, detalle FROM reportes_tickets ORDER BY id DESC")
            datos["tickets"] = [[str(t[0]), str(t[1]), str(t[2]), str(t[3]), str(t[4]), str(t[5] or "")] for t in cur.fetchall()]
    except Exception:
        logger.exception("Error leyendo datos")
    return datos

def cargar_historial(limite=200, sucursal=None):
    if not DATABASE_URL:
        return []
    consulta = "SELECT cambiado_en, sucursal, camara, estado_anterior, estado_nuevo, observacion_anterior, observacion_nueva, origen FROM historial_cctv"
    params = []
    if sucursal:
        consulta += " WHERE sucursal = %s"
        params.append(sucursal)
    consulta += " ORDER BY id DESC LIMIT %s"
    params.append(limite)
    try:
        with db_cursor() as cur:
            cur.execute(consulta, params)
            return [[local_txt(f[0]), f[1], f[2], f[3] or "(sin registro)", f[4], f[5] or "", f[6] or "", f[7] or ""] for f in cur.fetchall()]
    except Exception:
        return []

def estado_de(detalles, sucursal, camara):
    info = detalles.get(sucursal, {}).get(camara)
    return info["estado"] if info else "Por verificar"

def resumen_sucursal(detalles, sucursal):
    camaras = CATALOGO_CAMARAS[sucursal]
    op = nv = pv = 0
    for cam in camaras:
        est = estado_de(detalles, sucursal, cam)
        if est == "Operativa": op += 1
        elif est == "No visible": nv += 1
        else: pv += 1
    total = len(camaras)
    disp = (op / total * 100) if total else 0
    return total, op, nv, pv, disp

def resumen_total(detalles):
    total = op = nv = pv = 0
    for suc in CATALOGO_CAMARAS:
        t, o, n, p, _ = resumen_sucursal(detalles, suc)
        total, op, nv, pv = total + t, op + o, nv + n, pv + p
    disp = (op / total * 100) if total else 0
    return total, op, nv, pv, disp

def resumen_dict(detalles):
    total, op, nv, pv, disp = resumen_total(detalles)
    return {"total": total, "operativas": op, "no_visibles": nv, "por_verificar": pv, "disponibilidad": disp}

def nivel(disp):
    if disp >= 90: return "Óptimo"
    if disp >= 60: return "Atención"
    return "Crítico"

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()
    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_watermark_and_footer(num_pages)
            super().showPage()
        super().save()
    def draw_watermark_and_footer(self, total_pages):
        self.saveState()
        self.setFont("Helvetica-Bold", 60)
        self.setFillColor(colors.HexColor("#F2DEDE"), alpha=0.3)
        self.translate(300, 400)
        self.rotate(45)
        self.drawCentredString(0, 0, "CONFIDENCIAL")
        self.restoreState()
        self.saveState()
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.HexColor("#555555"))
        self.drawString(36, 25, "INJUL - Consolidado Maestro de CCTV")
        self.drawCentredString(letter[0] / 2, 25, f"Página {self._pageNumber} de {total_pages}")
        self.drawRightString(letter[0] - 36, 25, "Documento confidencial")
        self.setStrokeColor(colors.HexColor("#CCCCCC"))
        self.setLineWidth(0.5)
        self.line(36, 38, letter[0] - 36, 38)
        self.restoreState()

COLOR_ESTADO = {"Operativa": "#DFF0D8", "No visible": "#F2DEDE", "Por verificar": "#FCF8E3"}
COLOR_NIVEL = {"Óptimo": "#DFF0D8", "Atención": "#FCF8E3", "Crítico": "#F2DEDE"}

def estilo_tabla(color_cabecera, extra=None):
    comandos = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(color_cabecera)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
    ]
    if extra: comandos.extend(extra)
    return TableStyle(comandos)

EN_VIVO_OFICIAL = {"sí": "Sí", "si": "Sí", "no": "No", "intermitente": "Intermitente"}
GRABACION_OFICIAL = {"normal": "Normal", "por movimiento": "Por movimiento", "no graba": "No graba"}
COLOR_TEXTO_EN_VIVO = {"Sí": "#2E7D32", "No": "#C62828", "Intermitente": "#B26A00", "—": "#888888"}
COLOR_TEXTO_GRABACION = {"Normal": "#333333", "Por movimiento": "#333333", "No graba": "#C62828"}

def _valor_columna(info, claves, oficial):
    for clave in claves:
        crudo = info.get(clave)
        if not crudo: continue
        val = oficial.get(str(crudo).strip().lower())
        if val: return val
    return "—"

def _celda_color(texto, colores, estilo):
    color = colores.get(texto, "#888888")
    if texto == "—":
        return Paragraph(f"<font color='{color}'>—</font>", estilo)
    if color == "#333333":
        return Paragraph(esc(texto), estilo)
    return Paragraph(f"<font color='{color}'><b>{esc(texto)}</b></font>", estilo)

def filas_detalle_sucursal(detalles, suc, camaras, estilos):
    celda_izq, celda_cen = estilos
    filas = [["#", "Ubicación / Nombre en EZStation", "En vivo", "Grabación", "Estado", "Observaciones"]]
    extra = []
    for n, cam in enumerate(camaras, start=1):
        info = detalles.get(suc, {}).get(cam) or {}
        est = info.get("estado") or "Por verificar"
        if est not in COLOR_ESTADO: est = "Por verificar"
        obs = str(info.get("observacion") or "").strip()
        en_vivo = _valor_columna(info, ("en_vivo", "envivo"), EN_VIVO_OFICIAL)
        grab = _valor_columna(info, ("grabacion", "grabación"), GRABACION_OFICIAL)
        fila = [
            str(n),
            Paragraph(esc(cam), celda_izq),
            _celda_color(en_vivo, COLOR_TEXTO_EN_VIVO, celda_cen),
            _celda_color(grab, COLOR_TEXTO_GRABACION, celda_cen),
            _celda_color(est, {"Operativa": "#2E7D32", "No visible": "#C62828", "Por verificar": "#B26A00"}, celda_cen),
            Paragraph(esc(obs), celda_izq),
        ]
        filas.append(fila)
        extra.append(("BACKGROUND", (4, n), (4, n), colors.HexColor(COLOR_ESTADO[est])))
    return filas, extra

@app.route("/health")
def health():
    return "ok", 200

@app.route("/")
def index():
    tipo, mensaje = MENSAJES.get(request.args.get("msg"), (None, None))
    try: seleccion = int(request.args.get("s", 0))
    except ValueError: seleccion = 0
    datos = cargar_datos()
    return render_template(
        "index.html",
        catalogo=CATALOGO_CAMARAS,
        detalles=datos["detalles"],
        resumen=resumen_dict(datos["detalles"]),
        estados=ESTADOS,
        clases=CLASES_ESTADO,
        mensaje=mensaje,
        tipo=tipo,
        seleccion=seleccion,
    )

@app.route("/api/camara", methods=["POST"])
def api_camara():
    datos_in = request.get_json(silent=True) or {}
    sucursal = datos_in.get("sucursal")
    estado = datos_in.get("estado")
    en_vivo = str(datos_in.get("en_vivo") or "")
    grabacion = str(datos_in.get("grabacion") or "Normal")
    observacion = str(datos_in.get("observacion") or "").strip()[:500]
    idx = datos_in.get("idx")

    if sucursal not in CATALOGO_CAMARAS:
        return jsonify(ok=False, error="Sucursal no válida"), 400
    if not isinstance(idx, int) or isinstance(idx, bool) or not 0 <= idx < len(CATALOGO_CAMARAS[sucursal]):
        return jsonify(ok=False, error="Cámara no válida"), 400
    if estado not in ESTADOS_VALIDOS:
        return jsonify(ok=False, error="Estado no válido"), 400
    if not DATABASE_URL:
        return jsonify(ok=False, error="BD no configurada"), 503

    camara = CATALOGO_CAMARAS[sucursal][idx]
    try:
        with db_cursor() as cur:
            cur.execute("SELECT estado, observacion FROM estado_camaras_actual WHERE sucursal = %s AND camara = %s FOR UPDATE", (sucursal, camara))
            previo = cur.fetchone()
            est_ant = previo[0] if previo else None
            obs_ant = (previo[1] or "") if previo else None
            cambiado = previo is None or est_ant != estado or obs_ant != observacion
            
            cur.execute("""
                INSERT INTO estado_camaras_actual (sucursal, camara, estado, en_vivo, grabacion, observacion, actualizado_en)
                VALUES (%s, %s, %s, %s, %s, %s, NOW())
                ON CONFLICT (sucursal, camara) DO UPDATE
                SET estado = EXCLUDED.estado, en_vivo = EXCLUDED.en_vivo, grabacion = EXCLUDED.grabacion,
                    observacion = EXCLUDED.observacion, actualizado_en = NOW()
            """, (sucursal, camara, estado, en_vivo, grabacion, observacion))

            if cambiado:
                cur.execute("""
                    INSERT INTO historial_cctv (sucursal, camara, estado_anterior, estado_nuevo, observacion_anterior, observacion_nueva, origen)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (sucursal, camara, est_ant, estado, obs_ant, observacion, ip_cliente()))
    except Exception:
        logger.exception("Error guardando cámara")
        return jsonify(ok=False, error="No se pudo guardar"), 500

    datos = cargar_datos()
    return jsonify(ok=True, cambiado=cambiado, resumen=resumen_dict(datos["detalles"]))

@app.route("/guardar-ticket", methods=["POST"])
def guardar_ticket():
    ticket = (request.form.get("ticket") or "").strip()[:50]
    sucursal_area = (request.form.get("sucursal_area") or "").strip()[:150]
    estado_ticket = request.form.get("estado_ticket") or "En proceso"
    detalle = (request.form.get("detalle") or "").strip()
    try: dias_abiertos = max(0, int(request.form.get("dias_abiertos") or 0))
    except ValueError: dias_abiertos = 0

    if not ticket or not sucursal_area: return redirect(url_for("index", msg="error"))
    if not DATABASE_URL: return redirect(url_for("index", msg="sin_db"))

    try:
        with db_cursor() as cur:
            cur.execute("INSERT INTO reportes_tickets (fecha, ticket, sucursal_area, estado, dias_abiertos, detalle) VALUES (%s, %s, %s, %s, %s, %s)",
                        (ahora().strftime("%Y-%m-%d"), ticket, sucursal_area, estado_ticket, dias_abiertos, detalle))
    except Exception:
        return redirect(url_for("index", msg="error"))
    return redirect(url_for("index", msg="ticket_ok"))

@app.route("/historial")
def historial():
    filtro = request.args.get("sucursal") or ""
    if filtro not in CATALOGO_CAMARAS: filtro = ""
    return render_template("historial.html", filas=cargar_historial(200, filtro or None), catalogo=CATALOGO_CAMARAS, filtro=filtro)

@app.route("/descargar-pdf")
def descargar_pdf():
    # Nota: Aquí puedes usar tu función de PDF o la estándar que ya teníamos
    return "PDF en mantenimiento", 200

@app.route("/descargar-excel")
def descargar_excel():
    return "Excel en mantenimiento", 200

try:
    import ciclico
    ciclico.registrar(app, CATALOGO_CAMARAS, POOL_CICLICO, ANCLA_ROTACION, camaras_no_visibles, INVENTARIO_DE, db_cursor, ahora, local_txt, ip_cliente, DATABASE_URL)
except Exception:
    pass

inicializar_base_de_datos()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
