"""INJUL - Consolidado Maestro de CCTV (Flask + Postgres + Redis opcional)."""

import json
import logging
import os
from contextlib import contextmanager
from datetime import datetime
from io import BytesIO
from xml.sax.saxutils import escape

import openpyxl
import psycopg2
import redis
from flask import Flask, redirect, render_template_string, request, send_file, url_for
from openpyxl.styles import Alignment, Font, PatternFill
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
except Exception:  # noqa: BLE001 - si falta tzdata se usa la hora del servidor
    ZONA_HORARIA = None

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("injul-cctv")

app = Flask(__name__)

DATABASE_URL = os.environ.get("DATABASE_URL")
REDIS_URL = os.environ.get("REDIS_URL")
CACHE_KEY = "cctv:datos_reporte"
CACHE_TTL = 300  # segundos

# --- REDIS (CACHÉ OPCIONAL: si falla, la app sigue funcionando sin caché) ---
cache = None
if REDIS_URL:
    try:
        cache = redis.Redis.from_url(
            REDIS_URL,
            socket_connect_timeout=2,
            socket_timeout=2,
            decode_responses=True,
        )
    except Exception:  # noqa: BLE001
        logger.exception("No se pudo configurar Redis; se continúa sin caché")
        cache = None

# --- CATÁLOGO OFICIAL DE CÁMARAS POR SUCURSAL ---
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
        "Escritorio Garit", "Frente Puerta no", "Almacen Repuesto (2)", "Ing y Calidad",
        "Alm Gene 07", "Area de Planta", "ARTECH 1", "ARTECH 2", "Alm Gene 05",
        "IP Camera 57", "IP Camera 58", "IP Camera 59",
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
ESTADOS = [
    ("Operativa", "Operativa (Sí)"),
    ("No visible", "No visible (Rojo)"),
    ("Por verificar", "Por verificar"),
]

MENSAJES = {
    "sucursal_ok": ("ok", "Estado de la sucursal guardado correctamente."),
    "ticket_ok": ("ok", "Ticket registrado correctamente."),
    "sin_db": ("error", "La base de datos no está configurada (falta DATABASE_URL)."),
    "error": ("error", "No se pudo guardar. Revisa los registros en Railway."),
}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>INJUL - Consolidado CCTV</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 15px; }
        .container { max-width: 950px; margin: 0 auto; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        h1 { color: #1F497D; font-size: 20px; border-bottom: 2px solid #1F497D; padding-bottom: 8px; }
        h3 { color: #1F497D; margin-top: 5px; }
        .card { background: #f8f9fa; padding: 15px; margin-bottom: 15px; border-radius: 6px; border-left: 4px solid #1F497D; }
        .btn { display: inline-block; background: #1F497D; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; font-weight: bold; margin-top: 5px; margin-right: 5px; border: none; cursor: pointer; font-size: 14px; }
        .btn-excel { background: #217346; }
        .btn-ticket { background: #D9534F; }
        .btn:hover { opacity: 0.9; }
        .form-group { margin-bottom: 10px; }
        label { display: block; font-weight: bold; font-size: 13px; margin-bottom: 3px; color: #444; }
        input, select, textarea { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; font-size: 12px; }
        th { background-color: #1F497D; color: white; }
        .camera-row { background: #fff; }
        .aviso { padding: 10px 14px; border-radius: 4px; margin-bottom: 15px; font-size: 14px; }
        .aviso.ok { background: #dff0d8; color: #2d6a2d; }
        .aviso.error { background: #f2dede; color: #a94442; }
        .tabla-scroll { overflow-x: auto; }
    </style>
    <script>
        function actualizarCamaras() {
            var select = document.getElementById('sucursal-select');
            var listas = document.getElementsByClassName('lista-sucursal');
            for (var i = 0; i < listas.length; i++) {
                var activa = (listas[i].id === 'sucursal-' + select.selectedIndex);
                listas[i].style.display = activa ? 'block' : 'none';
                var campos = listas[i].querySelectorAll('select, input');
                for (var j = 0; j < campos.length; j++) {
                    campos[j].disabled = !activa;
                }
            }
        }
        window.onload = actualizarCamaras;
    </script>
</head>
<body>
    <div class="container">
        <h1>INJUL - Sistema de Monitoreo CCTV</h1>

        {% if mensaje %}<div class="aviso {{ tipo }}">{{ mensaje }}</div>{% endif %}

        <div class="card">
            <h3>Seleccionar Sucursal y Revisar Cámaras</h3>
            <form action="/guardar-sucursal" method="POST">
                <div class="form-group">
                    <label>Sucursal:</label>
                    <select id="sucursal-select" name="sucursal" onchange="actualizarCamaras()">
                        {% for suc in catalogo.keys() %}
                        <option value="{{ suc }}" {% if loop.index0 == seleccion %}selected{% endif %}>{{ suc }}</option>
                        {% endfor %}
                    </select>
                </div>

                {% for suc, camaras in catalogo.items() %}
                {% set guardado = detalles.get(suc, {}) %}
                <div id="sucursal-{{ loop.index0 }}" class="lista-sucursal" style="display:none; margin-top: 15px;">
                    <h4 style="color: #1F497D; margin-bottom: 5px;">Cámaras de {{ suc }}</h4>
                    <p style="font-size: 12px; color: #666;">Selecciona el estado de cada cámara para el reporte:</p>
                    <div class="tabla-scroll">
                    <table>
                        <tr>
                            <th>#</th>
                            <th>Nombre de la Cámara (EZStation)</th>
                            <th>Estado</th>
                            <th>Observación</th>
                        </tr>
                        {% for cam in camaras %}
                        {% set info = guardado.get(cam, {}) %}
                        <tr class="camera-row">
                            <td>{{ loop.index }}</td>
                            <td style="text-align: left; font-weight: bold;">{{ cam }}</td>
                            <td>
                                <select name="estado_{{ loop.index0 }}" style="padding: 4px;">
                                    {% for valor, texto in estados %}
                                    <option value="{{ valor }}" {% if (info.get('estado') or 'Operativa') == valor %}selected{% endif %}>{{ texto }}</option>
                                    {% endfor %}
                                </select>
                            </td>
                            <td>
                                <input type="text" name="obs_{{ loop.index0 }}" value="{{ info.get('observacion', '') }}" placeholder="Detalle si falla..." style="padding: 4px;">
                            </td>
                        </tr>
                        {% endfor %}
                    </table>
                    </div>
                </div>
                {% endfor %}

                <div style="margin-top: 15px;">
                    <button type="submit" class="btn">💾 Guardar Estado de esta Sucursal</button>
                </div>
            </form>
        </div>

        <div class="card">
            <h3>Registrar Ticket o Novedad con Proveedor</h3>
            <form action="/guardar-ticket" method="POST">
                <div class="form-group">
                    <label>Número de Ticket:</label>
                    <input type="text" name="ticket" required maxlength="50" placeholder="Ej: 10437 o SIN TICKET">
                </div>
                <div class="form-group">
                    <label>Sucursal o Área Afectada:</label>
                    <input type="text" name="sucursal_area" required maxlength="150" placeholder="Ej: Marmotech DCM">
                </div>
                <div class="form-group">
                    <label>Estado del Ticket:</label>
                    <select name="estado_ticket">
                        <option value="En proceso">En proceso</option>
                        <option value="Terminado">Terminado</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Días Abiertos:</label>
                    <input type="number" name="dias_abiertos" value="0" min="0" placeholder="Ej: 5">
                </div>
                <div class="form-group">
                    <label>Detalle del Reporte (Información de lo que pasó):</label>
                    <textarea name="detalle" rows="2" placeholder="Ej: Problemas de visualización..."></textarea>
                </div>
                <button type="submit" class="btn btn-ticket">📝 Registrar Ticket</button>
            </form>
        </div>

        <div class="card">
            <h3>Generación de Reportes Oficiales</h3>
            <a href="/descargar-pdf" class="btn">📥 Descargar PDF Oficial</a>
            <a href="/descargar-excel" class="btn btn-excel">📊 Descargar Excel Matriz</a>
        </div>
    </div>
</body>
</html>
"""


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def ahora():
    return datetime.now(ZONA_HORARIA) if ZONA_HORARIA else datetime.now()


def esc(texto):
    """Escapa texto para usarlo dentro de un Paragraph de ReportLab."""
    return escape(str(texto if texto is not None else ""))


@contextmanager
def db_cursor():
    """Abre conexión a Postgres; confirma al salir bien y revierte si hay error."""
    conn = psycopg2.connect(DATABASE_URL, connect_timeout=5)
    try:
        with conn:
            with conn.cursor() as cur:
                yield cur
    finally:
        conn.close()


def inicializar_base_de_datos():
    if not DATABASE_URL:
        logger.warning("DATABASE_URL no está definida; no se guardarán datos")
        return
    try:
        with db_cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS reportes_cctv_detalles (
                    id SERIAL PRIMARY KEY,
                    fecha DATE NOT NULL,
                    sucursal VARCHAR(100) NOT NULL,
                    camara VARCHAR(150) NOT NULL,
                    estado VARCHAR(50) NOT NULL,
                    observacion TEXT
                );
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
            """)
    except Exception:  # noqa: BLE001
        logger.exception("Error inicializando la base de datos")


# ----------------------------------------------------------------------------
# Caché (Redis) y carga de datos
# ----------------------------------------------------------------------------
def cache_leer():
    if cache is None:
        return None
    try:
        crudo = cache.get(CACHE_KEY)
        return json.loads(crudo) if crudo else None
    except Exception:  # noqa: BLE001
        logger.warning("Redis no disponible al leer; se consulta la base de datos")
        return None


def cache_guardar(datos):
    if cache is None:
        return
    try:
        cache.setex(CACHE_KEY, CACHE_TTL, json.dumps(datos))
    except Exception:  # noqa: BLE001
        logger.warning("Redis no disponible al guardar caché")


def cache_invalidar():
    if cache is None:
        return
    try:
        cache.delete(CACHE_KEY)
    except Exception:  # noqa: BLE001
        logger.warning("Redis no disponible al invalidar caché")


def cargar_datos():
    """Devuelve {'detalles': {sucursal: {camara: {...}}}, 'tickets': [[...]]}."""
    datos = cache_leer()
    if datos is not None:
        return datos

    datos = {"detalles": {}, "tickets": []}
    if not DATABASE_URL:
        return datos
    try:
        with db_cursor() as cur:
            cur.execute("""
                SELECT sucursal, camara, estado, observacion
                FROM reportes_cctv_detalles
                WHERE fecha = (SELECT MAX(fecha) FROM reportes_cctv_detalles)
            """)
            for suc, cam, est, obs in cur.fetchall():
                datos["detalles"].setdefault(suc, {})[cam] = {
                    "estado": est,
                    "observacion": obs or "",
                }
            cur.execute("""
                SELECT ticket, fecha, sucursal_area, estado, dias_abiertos, detalle
                FROM reportes_tickets
                ORDER BY id DESC
            """)
            datos["tickets"] = [
                [str(t[0]), str(t[1]), str(t[2]), str(t[3]), str(t[4]), str(t[5] or "")]
                for t in cur.fetchall()
            ]
        cache_guardar(datos)
    except Exception:  # noqa: BLE001
        logger.exception("Error leyendo datos de la base de datos")
    return datos


def estado_de(detalles, sucursal, camara):
    info = detalles.get(sucursal, {}).get(camara)
    return info["estado"] if info else "Por verificar"


def resumen_sucursal(detalles, sucursal):
    camaras = CATALOGO_CAMARAS[sucursal]
    op = nv = pv = 0
    for cam in camaras:
        est = estado_de(detalles, sucursal, cam)
        if est == "Operativa":
            op += 1
        elif est == "No visible":
            nv += 1
        else:
            pv += 1
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


# ----------------------------------------------------------------------------
# PDF
# ----------------------------------------------------------------------------
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
        self.drawString(36, 25, "INJUL Consolidado Maestro de CCTV  |  Documento confidencial")
        self.drawRightString(letter[0] - 36, 25, f"Página {self._pageNumber} de {total_pages}")
        self.setStrokeColor(colors.HexColor("#CCCCCC"))
        self.setLineWidth(0.5)
        self.line(36, 38, letter[0] - 36, 38)
        self.restoreState()


COLOR_ESTADO = {
    "Operativa": "#DFF0D8",
    "No visible": "#F2DEDE",
    "Por verificar": "#FCF8E3",
}


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
    if extra:
        comandos.extend(extra)
    return TableStyle(comandos)


def construir_pdf(datos):
    detalles = datos["detalles"]
    tickets = datos["tickets"]
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=50,
    )
    title_style = ParagraphStyle(
        "TitleStyle",
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=19,
        textColor=colors.HexColor("#1F497D"),
    )
    normal_style = ParagraphStyle(
        "NormalStyle",
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#333333"),
    )
    celda = ParagraphStyle("Celda", fontName="Helvetica", fontSize=8, leading=10)
    story = []

    # --- PORTADA ---
    fecha_txt = ahora().strftime("%d/%m/%Y")
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
    encabezado = Paragraph(
        "<b>INJUL - Consolidado Maestro de CCTV</b><br/>"
        "<font size=9 color='#555555'>Investigadores Jurídicos EIRL<br/>"
        f"Fecha: {fecha_txt} | Rango: 07:00 AM - 07:00 PM</font>",
        title_style,
    )
    if os.path.exists(logo_path):
        logo_img = Image(logo_path, width=1.1 * inch, height=1.1 * inch)
        t_header = Table([[logo_img, encabezado]], colWidths=[85, 455])
        t_header.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        story.append(t_header)
    else:
        story.append(encabezado)
    story.append(Spacer(1, 15))

    total, op, nv, pv, disp = resumen_total(detalles)
    kpi_data = [
        ["Cámaras Totales", "Operativas", "No Visibles", "Por Verificar", "Disponibilidad"],
        [str(total), str(op), str(nv), str(pv), f"{disp:.1f}%"],
    ]
    t_kpi = Table(kpi_data, colWidths=[100, 100, 100, 100, 140])
    t_kpi.setStyle(
        estilo_tabla("#1F497D", [("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F2F2F2"))])
    )
    story.append(t_kpi)
    story.append(PageBreak())

    # --- HALLAZGOS Y RESUMEN ---
    story.append(Paragraph("<b>Hallazgos del Día y Resumen de Tickets</b>", title_style))
    story.append(Spacer(1, 10))
    en_proceso = sum(1 for t in tickets if t[3] == "En proceso")
    story.append(
        Paragraph(
            "1. Monitoreo activo de la red de sucursales en EZStation.<br/>"
            "2. Seguimiento de incidencias técnicas detalladas por cámara.<br/>"
            f"3. Tickets registrados: {len(tickets)} (en proceso: {en_proceso}).",
            normal_style,
        )
    )
    story.append(Spacer(1, 12))
    filas = [["Sucursal", "Total", "Operativas", "No visibles", "Por verificar", "Disponib."]]
    for suc in CATALOGO_CAMARAS:
        t, o, n, p, d = resumen_sucursal(detalles, suc)
        filas.append([Paragraph(esc(suc), celda), str(t), str(o), str(n), str(p), f"{d:.1f}%"])
    t_res = Table(filas, colWidths=[190, 50, 70, 70, 80, 80], repeatRows=1)
    t_res.setStyle(estilo_tabla("#1F497D"))
    story.append(t_res)
    story.append(PageBreak())

    # --- UNA SECCIÓN POR SUCURSAL ---
    total_suc = len(CATALOGO_CAMARAS)
    for idx, (suc, camaras) in enumerate(CATALOGO_CAMARAS.items(), start=1):
        story.append(Paragraph(f"<b>Sucursal {idx} de {total_suc}: {esc(suc)}</b>", title_style))
        story.append(Spacer(1, 10))
        filas = [["#", "Ubicación / Nombre en EZStation", "En vivo", "Grabación", "Estado", "Observaciones"]]
        extra = [("ALIGN", (1, 1), (1, -1), "LEFT"), ("ALIGN", (5, 1), (5, -1), "LEFT")]
        for n, cam in enumerate(camaras, start=1):
            est = estado_de(detalles, suc, cam)
            obs = detalles.get(suc, {}).get(cam, {}).get("observacion", "")
            en_vivo = {"Operativa": "Sí", "No visible": "No"}.get(est, "Por verificar")
            grab = {"Operativa": "Normal", "No visible": "Revisar"}.get(est, "Por verificar")
            filas.append([str(n), Paragraph(esc(cam), celda), en_vivo, grab, est, Paragraph(esc(obs), celda)])
            extra.append(("BACKGROUND", (4, n), (4, n), colors.HexColor(COLOR_ESTADO.get(est, "#FCF8E3"))))
        t_suc = Table(filas, colWidths=[25, 185, 50, 70, 70, 140], repeatRows=1)
        t_suc.setStyle(estilo_tabla("#333333", extra))
        story.append(t_suc)
        story.append(PageBreak())

    # --- TICKETS ---
    story.append(Paragraph("<b>Seguimiento de Tickets con el Proveedor</b>", title_style))
    story.append(Spacer(1, 10))
    filas = [["Ticket", "Fecha", "Sucursal / Área", "Estado", "Días", "Detalle"]]
    if tickets:
        for t in tickets:
            filas.append([
                Paragraph(esc(t[0]), celda),
                t[1],
                Paragraph(esc(t[2]), celda),
                t[3],
                t[4],
                Paragraph(esc(t[5]), celda),
            ])
    else:
        filas.append(["N/A", "N/A", "Sin tickets registrados", "N/A", "0", "N/A"])
    t_tickets = Table(filas, colWidths=[60, 70, 120, 70, 40, 170], repeatRows=1)
    t_tickets.setStyle(estilo_tabla("#1F497D"))
    story.append(t_tickets)

    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
def construir_excel(datos):
    detalles = datos["detalles"]
    tickets = datos["tickets"]
    azul = PatternFill("solid", fgColor="1F497D")
    blanco = Font(bold=True, color="FFFFFF")
    rellenos = {
        "Operativa": PatternFill("solid", fgColor="DFF0D8"),
        "No visible": PatternFill("solid", fgColor="F2DEDE"),
        "Por verificar": PatternFill("solid", fgColor="FCF8E3"),
    }

    def cabecera(ws, fila, titulos):
        for col, titulo in enumerate(titulos, start=1):
            c = ws.cell(row=fila, column=col, value=titulo)
            c.fill, c.font = azul, blanco
            c.alignment = Alignment(horizontal="center", vertical="center")

    wb = openpyxl.Workbook()

    # Hoja 1: resumen
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = "INJUL - Consolidado de CCTV"
    ws["A1"].font = Font(size=14, bold=True, color="1F497D")
    ws["A2"] = f"Fecha: {ahora().strftime('%d/%m/%Y')}"
    cabecera(ws, 4, ["Sucursal", "Total", "Operativas", "No visibles", "Por verificar", "Disponibilidad %"])
    fila = 5
    for suc in CATALOGO_CAMARAS:
        t, o, n, p, d = resumen_sucursal(detalles, suc)
        for col, valor in enumerate([suc, t, o, n, p, round(d, 1)], start=1):
            ws.cell(row=fila, column=col, value=valor)
        fila += 1
    t, o, n, p, d = resumen_total(detalles)
    for col, valor in enumerate(["TOTAL", t, o, n, p, round(d, 1)], start=1):
        ws.cell(row=fila, column=col, value=valor).font = Font(bold=True)
    for col, ancho in zip("ABCDEF", [34, 10, 12, 13, 14, 18]):
        ws.column_dimensions[col].width = ancho

    # Hoja 2: matriz de cámaras
    wm = wb.create_sheet("Matriz")
    cabecera(wm, 1, ["Sucursal", "#", "Cámara (EZStation)", "Estado", "Observación"])
    fila = 2
    for suc, camaras in CATALOGO_CAMARAS.items():
        for n, cam in enumerate(camaras, start=1):
            est = estado_de(detalles, suc, cam)
            obs = detalles.get(suc, {}).get(cam, {}).get("observacion", "")
            for col, valor in enumerate([suc, n, cam, est, obs], start=1):
                wm.cell(row=fila, column=col, value=valor)
            wm.cell(row=fila, column=4).fill = rellenos.get(est, rellenos["Por verificar"])
            fila += 1
    wm.freeze_panes = "A2"
    for col, ancho in zip("ABCDE", [30, 6, 42, 16, 50]):
        wm.column_dimensions[col].width = ancho

    # Hoja 3: tickets
    wt = wb.create_sheet("Tickets")
    cabecera(wt, 1, ["Ticket", "Fecha", "Sucursal / Área", "Estado", "Días abiertos", "Detalle"])
    for fila, t in enumerate(tickets, start=2):
        for col, valor in enumerate(t, start=1):
            wt.cell(row=fila, column=col, value=valor)
    wt.freeze_panes = "A2"
    for col, ancho in zip("ABCDEF", [14, 14, 30, 14, 14, 60]):
        wt.column_dimensions[col].width = ancho

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------------------------------
# Rutas
# ----------------------------------------------------------------------------
@app.route("/health")
def health():
    return "ok", 200


@app.route("/")
def index():
    tipo, mensaje = MENSAJES.get(request.args.get("msg"), (None, None))
    try:
        seleccion = int(request.args.get("s", 0))
    except ValueError:
        seleccion = 0
    return render_template_string(
        HTML_TEMPLATE,
        catalogo=CATALOGO_CAMARAS,
        detalles=cargar_datos()["detalles"],
        estados=ESTADOS,
        mensaje=mensaje,
        tipo=tipo,
        seleccion=seleccion,
    )


@app.route("/guardar-sucursal", methods=["POST"])
def guardar_sucursal():
    sucursal = request.form.get("sucursal")
    if sucursal not in CATALOGO_CAMARAS:
        return redirect(url_for("index", msg="error"))
    seleccion = list(CATALOGO_CAMARAS).index(sucursal)
    if not DATABASE_URL:
        return redirect(url_for("index", msg="sin_db", s=seleccion))

    fecha_hoy = ahora().strftime("%Y-%m-%d")
    validos = {valor for valor, _ in ESTADOS}
    filas = []
    for i, cam in enumerate(CATALOGO_CAMARAS[sucursal]):
        estado = request.form.get(f"estado_{i}") or "Operativa"
        if estado not in validos:
            estado = "Por verificar"
        obs = (request.form.get(f"obs_{i}") or "").strip()
        filas.append((fecha_hoy, sucursal, cam, estado, obs))

    try:
        with db_cursor() as cur:
            cur.execute(
                "DELETE FROM reportes_cctv_detalles WHERE fecha = %s AND sucursal = %s",
                (fecha_hoy, sucursal),
            )
            cur.executemany(
                """
                INSERT INTO reportes_cctv_detalles (fecha, sucursal, camara, estado, observacion)
                VALUES (%s, %s, %s, %s, %s)
                """,
                filas,
            )
    except Exception:  # noqa: BLE001
        logger.exception("Error guardando sucursal")
        return redirect(url_for("index", msg="error", s=seleccion))

    cache_invalidar()
    return redirect(url_for("index", msg="sucursal_ok", s=seleccion))


@app.route("/guardar-ticket", methods=["POST"])
def guardar_ticket():
    ticket = (request.form.get("ticket") or "").strip()[:50]
    sucursal_area = (request.form.get("sucursal_area") or "").strip()[:150]
    estado_ticket = request.form.get("estado_ticket") or "En proceso"
    if estado_ticket not in ("En proceso", "Terminado"):
        estado_ticket = "En proceso"
    detalle = (request.form.get("detalle") or "").strip()
    try:
        dias_abiertos = max(0, int(request.form.get("dias_abiertos") or 0))
    except ValueError:
        dias_abiertos = 0

    if not ticket or not sucursal_area:
        return redirect(url_for("index", msg="error"))
    if not DATABASE_URL:
        return redirect(url_for("index", msg="sin_db"))

    try:
        with db_cursor() as cur:
            cur.execute(
                """
                INSERT INTO reportes_tickets
                    (fecha, ticket, sucursal_area, estado, dias_abiertos, detalle)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (ahora().strftime("%Y-%m-%d"), ticket, sucursal_area, estado_ticket, dias_abiertos, detalle),
            )
    except Exception:  # noqa: BLE001
        logger.exception("Error guardando ticket")
        return redirect(url_for("index", msg="error"))

    cache_invalidar()
    return redirect(url_for("index", msg="ticket_ok"))


@app.route("/descargar-pdf")
def descargar_pdf():
    buffer = construir_pdf(cargar_datos())
    return send_file(
        buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="Reporte_CCTV_INJUL.pdf",
    )


@app.route("/descargar-excel")
def descargar_excel():
    buffer = construir_excel(cargar_datos())
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Reporte_CCTV_INJUL.xlsx",
    )


# Se ejecuta al importar, para que funcione tanto con gunicorn como con python main.py
inicializar_base_de_datos()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
