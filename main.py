import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import psycopg2
import redis
from flask import Flask, redirect, render_template_string, request, send_file, url_for
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
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

app = Flask(__name__)

# --- CONFIGURACIÓN DE REDIS (CACHÉ) ---
# Lee automáticamente la variable de entorno REDIS_URL configurada en Railway
redis_url = os.environ.get("REDIS_URL", "redis://localhost:6379")
cache = redis.Redis.from_url(redis_url)

# --- CATÁLOGO OFICIAL DE CÁMARAS POR SUCURSAL ---
CATALOGO_CAMARAS = {
    "Caribbean Agregados": [
        "Frente Oficina",
        "Puerta desde CR",
        "Lado Atras",
        "IP Camera 04",
        "Frente Lado desde Caribbean",
        "Lado Atras esquina puerta",
        "Puerta lado norte",
        "Frente hacia lado norte",
        "Oficina Gabinete data",
        "IP Camera 10",
        "IP Camera 11",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech Blue Mall": [
        "Show 2da planta",
        "Area Clientes 1ra planta",
        "Area 1ra planta puerta trasera",
        "2da Planta banos",
        "Area 1ra planta puerta trasera (2)",
        "Gabinete TI",
        "IP Camera 07",
        "2da Planta banos (2)",
        "IP Camera 09",
        "IP Camera 10",
        "IP Camera 11",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech DCM": [
        "Puerta Principal",
        "Cocina",
        "IP Camera 03",
        "S1 Entrada Vehiculos",
        "S2 Pasillo lateral",
        "IP Camera 06",
        "1ra planta asensor visita",
        "IP Camera 09",
        "IP Camera 11",
        "segundo nivel banera",
        "seg piso griferia",
        "IP Camera 14",
        "IP Camera 15",
        "Entrada edificio",
        "S2 Salida",
        "IP Camera 18",
        "IP Camera 19",
        "5to piso onix miel",
        "S2 Salida V",
        "S2 Salida Planta E",
        "Entrada Calle Parqueos",
        "Frente Edificio",
        "Salida S1 Vehiculos",
        "S1 Entrada al S2",
        "S1 Entrada vista de frente",
        "S2 Area Piezas y Vehiculos",
        "IP Camera 30",
        "almacen griferia piiso 1",
        "piso 1 entrada plancha",
        "IP Camera 33",
        "IP Camera 34",
        "IP Camera 35",
        "IP Camera 36",
        "IP Camera 37",
        "IP Camera 38",
        "IP Camera 39",
        "IP Camera 40",
        "IP Camera 41",
        "IP Camera 42",
        "IP Camera 43",
        "IP Camera 45",
        "Oficina Almacen",
        "Entrada a cobro lobby",
        "IP Camera 49",
        "Salida Asensor c",
        "IP Camera 51",
        "IP Camera 52",
        "IP Camera 53",
        "IP Camera 54",
        "IP Camera 55",
        "IP Camera 56",
    ],
    "Marmotech Outlet RP": [
        "IP Camera 01",
        "IP Camera 02",
        "IP Camera 03",
        "IP Camera 04",
        "IP Camera 05",
        "IP Camera 06",
        "IP Camera 07",
        "IP Camera 08",
        "IP Camera 09",
        "IP Camera 10",
        "IP Camera 11",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
    "Marmotech Punta Cana": [
        "Puerta Entrada Principal",
        "Area trasera 1",
        "Area Trasera 2",
        "Area Trasera 5",
        "Area trasera 3",
        "Area Trasera 4",
        "Oficina 1",
        "Area Gabinete datos",
        "Almacen",
        "Area Lateral Derecha",
        "Zona Carga / Descarga",
        "Almacen 1",
        "Frente Edificacion",
        "Instalacion",
        "Entrada Oficinas",
        "Almacen 2 Escalera Subir Techo Oficina",
        "Almacen 3",
    ],
    "Sucursal San Cristóbal": [
        "Almacen Repuesto",
        "Parqueo 1 Tienda",
        "Entrada Tienda S",
        "Parqueo 2 Tienda",
        "Almacen 1 Tienda",
        "C/Entrada Tienda",
        "Almacen 2 Tienda",
        "Almacen 3 Tienda",
        "Calle Entrada 2",
        "Oficina Tienda S",
        "Parqueo 3 Tiend",
        "Almacen 4 Tienda",
        "Almacen 5 Tienda",
        "Area Caja Oficin",
        "MODULARES 1",
        "Entrada Mantenimiento norte",
        "MODULARES Area F",
        "MODULARES 4",
        "Alm. Gene 04",
        "Alm Gene 06",
        "Cuarto de Compresores",
        "Alm Gene 01",
        "Alm Gene 02",
        "MOTORES 2",
        "ENTRADA P1",
        "Puerta Entrada",
        "ARTECH 02",
        "PRODUCCION 1",
        "ARTECH 01",
        "Alm Gene 03",
        "PULIDORAS",
        "ARTECH03",
        "PULIDORA 02",
        "DATA CENTER 1",
        "Carga Area 1 Alm",
        "Puerta Entrada A",
        "Area Motores Alm",
        "Carga 2 Alm Gene",
        "Gasoil entrada /",
        "Area Parqueo 1 F",
        "Salida Fabrica A",
        "Gasoil Dispensad",
        "Gasoil",
        "Parqueo Fabrica",
        "Puerta Entrada G",
        "Garita Guardian",
        "Escritorio Garit",
        "Frente Puerta no",
        "Almacen Repuesto (2)",
        "Ing y Calidad",
        "Alm Gene 07",
        "Area de Planta",
        "ARTECH 1",
        "ARTECH 2",
        "Alm Gene 05",
        "IP Camera 57",
        "IP Camera 58",
        "IP Camera 59",
    ],
    "Marmotech Santiago 1": [
        "Almacen Repuesto",
        "Parqueo 1 Tienda",
        "Entrada Tienda 5",
        "Parqueo 2 Tienda",
        "Almacen 1 Tienda",
        "C/Entrada Tienda",
        "Almacen 2 Tienda",
        "Almacen 3 Tienda",
        "Calle Entrada 2",
        "Oficina Tienda S",
        "Parqueo 3 Tiend",
        "Almacen 4 Tienda",
        "Almacen 5 Tienda",
        "Area Caja Oficin",
        "MODULARES 1",
        "Entrada Mantenimiento norte",
    ],
    "Marmotech Santiago 2": [
        "IP Camera 01",
        "IP Camera 02",
        "IP Camera 03",
        "IP Camera 04",
        "IP Camera 05",
        "IP Camera 06",
        "IP Camera 07",
        "IP Camera 08",
        "IP Camera 09",
        "IP Camera 10",
        "IP Camera 11",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
    "Santiago 2 - Sección Muebles": [
        "MUEBLES",
        "ENTRADA",
        "COCINA",
        "GRIFERIA",
        "PARQUEO O-E",
        "BANOS",
        "PISOS",
        "SALA REUNION",
        "PARQUEO E-O",
        "COUNTER",
        "MASETEROS",
        "IP Camera 12",
        "IP Camera 13",
        "IP Camera 14",
        "IP Camera 15",
        "IP Camera 16",
    ],
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
    </style>
    <script>
        function actualizarCamaras() {
            var select = document.getElementById('sucursal-select');
            var sucursalSeleccionada = select.value;
            var divs = document.getElementsByClassName('lista-sucursal');
            for (var i = 0; i < divs.length; i++) {
                divs[i].style.display = 'none';
            }
            var activo = document.getElementById('sucursal-' + sucursalSeleccionada);
            if (activo) {
                activo.style.display = 'block';
            }
            var selectedOption = select.options[select.selectedIndex];
            document.getElementById('nombre-sucursal').value = selectedOption.text;
        }
        window.onload = function() {
            actualizarCamaras();
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>INJUL - Sistema de Monitoreo CCTV</h1>
        
        <div class="card">
            <h3>Seleccionar Sucursal y Revisar Cámaras</h3>
            <div class="form-group">
                <label>Sucursal:</label>
                <select id="sucursal-select" onchange="actualizarCamaras()">
                    {% for suc in catalogo.keys() %}
                    <option value="{{ loop.index0 }}">{{ suc }}</option>
                    {% endfor %}
                </select>
            </div>

            <form action="/guardar-sucursal" method="POST">
                <input type="hidden" id="nombre-sucursal" name="sucursal" value="">
                
                {% for suc, camaras in catalogo.items() %}
                <div id="sucursal-{{ loop.index0 }}" class="lista-sucursal" style="display:none; margin-top: 15px;">
                    <h4 style="color: #1F497D; margin-bottom: 5px;">Cámaras de {{ suc }}</h4>
                    <p style="font-size: 12px; color: #666;">Selecciona el estado de cada cámara para el reporte:</p>
                    <table>
                        <tr>
                            <th>#</th>
                            <th>Nombre de la Cámara (EZStation)</th>
                            <th>Estado</th>
                            <th>Observación</th>
                        </tr>
                        {% for cam in camaras %}
                        <tr class="camera-row">
                            <td>{{ loop.index }}</td>
                            <td style="text-align: left; font-weight: bold;">{{ cam }}</td>
                            <td>
                                <select name="estado_{{ loop.index0 }}" style="padding: 4px;">
                                    <option value="Operativa">Operativa (Sí)</option>
                                    <option value="No visible">No visible (Rojo)</option>
                                    <option value="Por verificar">Por verificar</option>
                                </select>
                            </td>
                            <td>
                                <input type="text" name="obs_{{ loop.index0 }}" placeholder="Detalle si falla..." style="padding: 4px;">
                            </td>
                        </tr>
                        {% endfor %}
                    </table>
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
                    <input type="text" name="ticket" required placeholder="Ej: 10437 o SIN TICKET">
                </div>
                <div class="form-group">
                    <label>Sucursal o Área Afectada:</label>
                    <input type="text" name="sucursal_area" required placeholder="Ej: Marmotech DCM">
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
                    <input type="number" name="dias_abiertos" value="0" placeholder="Ej: 5">
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
            <a href="/descargar-pdf" class="btn">📥 Descargar PDF Oficial (13 Págs)</a>
            <a href="/descargar-excel" class="btn btn-excel">📊 Descargar Excel Matriz</a>
        </div>
    </div>
</body>
</html>
"""


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
    self.drawString(
        36,
        25,
        "INJUL Consolidado Maestro de CCTV                         Documento"
        " confidencial",
    )
    self.drawRightString(letter[0] - 36, 25, f"Página {self._pageNumber} de {total_pages}")
    self.setStrokeColor(colors.HexColor("#CCCCCC"))
    self.setLineWidth(0.5)
    self.line(36, 38, letter[0] - 36, 38)
    self.restoreState()


def inicializar_base_de_datos():
  database_url = os.environ.get("DATABASE_URL")
  if not database_url:
    return
  try:
    conn = psycopg2.connect(database_url)
    cur = conn.cursor()
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
    conn.commit()
    cur.close()
    conn.close()
  except Exception as e:
    print(f"Error DB: {e}")


@app.route("/")
def index():
  return render_template_string(HTML_TEMPLATE, catalogo=CATALOGO_CAMARAS)


@app.route("/guardar-sucursal", methods=["POST"])
def guardar_sucursal():
  sucursal = request.form.get("sucursal")
  fecha_hoy = datetime.now().strftime("%Y-%m-%d")

  database_url = os.environ.get("DATABASE_URL")
  if database_url and sucursal in CATALOGO_CAMARAS:
    try:
      conn = psycopg2.connect(database_url)
      cur = conn.cursor()
      cur.execute(
          "DELETE FROM reportes_cctv_detalles WHERE fecha = %s AND sucursal = %s",
          (fecha_hoy, sucursal),
      )

      camaras = CATALOGO_CAMARAS[sucursal]
      for i, cam in enumerate(camaras):
        estado = request.form.get(f"estado_{i}") or "Operativa"
        obs = request.form.get(f"obs_{i}") or ""
        cur.execute(
            """
                    INSERT INTO reportes_cctv_detalles (fecha, sucursal, camara, estado, observacion)
                    VALUES (%s, %s, %s, %s, %s)
                """,
            (fecha_hoy, sucursal, cam, estado, obs),
        )

      conn.commit()
      cur.close()
      conn.close()
      
      # Limpiamos caché de Redis al actualizar información
      try:
        cache.flushall()
      except Exception:
        pass

    except Exception as e:
      print(f"Error guardando sucursal: {e}")

  return redirect(url_for("index"))


@app.route("/guardar-ticket", methods=["POST"])
def guardar_ticket():
  ticket = request.form.get("ticket")
  sucursal_area = request.form.get("sucursal_area")
  estado_ticket = request.form.get("estado_ticket")
  dias_abiertos = request.form.get("dias_abiertos") or 0
  detalle = request.form.get("detalle")
  fecha_hoy = datetime.now().strftime("%Y-%m-%d")

  database_url = os.environ.get("DATABASE_URL")
  if database_url:
    try:
      conn = psycopg2.connect(database_url)
      cur = conn.cursor()
      cur.execute(
          """
                INSERT INTO reportes_tickets (fecha, ticket, sucursal_area, estado, dias_abiertos, detalle)
                VALUES (%s, %s, %s, %s, %s, %s)
            """,
          (
              fecha_hoy,
              ticket,
              sucursal_area,
              estado_ticket,
              int(dias_abiertos),
              detalle,
          ),
      )
      conn.commit()
      cur.close()
      conn.close()
      
      try:
        cache.flushall()
      except Exception:
        pass

    except Exception as e:
      print(f"Error guardando ticket: {e}")

  return redirect(url_for("index"))


@app.route("/descargar-pdf")
def descargar_pdf():
  pdf_path = "Reporte_CCTV_INJUL.pdf"
  doc = SimpleDocTemplate(
      pdf_path,
      pagesize=letter,
      leftMargin=36,
      rightMargin=36,
      topMargin=36,
      bottomMargin=50,
  )
  story = []
  from reportlab.lib.styles import ParagraphStyle

  title_style = ParagraphStyle(
      "TitleStyle",
      fontName="Helvetica-Bold",
      fontSize=15,
      textColor=colors.HexColor("#1F497D"),
  )
  normal_style = ParagraphStyle(
      "NormalStyle",
      fontName="Helvetica",
      fontSize=10,
      textColor=colors.HexColor("#333333"),
  )

  datos_detalles = {}
  lista_tickets_db = []
  database_url = os.environ.get("DATABASE_URL")
  if database_url:
    try:
      conn = psycopg2.connect(database_url)
      cur = conn.cursor()
      cur.execute("""
                SELECT sucursal, camara, estado, observacion 
                FROM reportes_cctv_detalles 
                WHERE fecha = (SELECT MAX(fecha) FROM reportes_cctv_detalles);
            """)
      for r in cur.fetchall():
        suc, cam, est, obs = r[0], r[1], r[2], r[3]
        if suc not in datos_detalles:
          datos_detalles[suc] = {}
        datos_detalles[suc][cam] = {"estado": est, "observacion": obs}

      cur.execute("""
                SELECT ticket, fecha, sucursal_area, estado, dias_abiertos, detalle
                FROM reportes_tickets
                ORDER BY id DESC;
            """)
      lista_tickets_db = cur.fetchall()
      cur.close()
      conn.close()
    except Exception as e:
      print(f"Error DB PDF: {e}")

  # --- PÁGINA 1: PORTADA ---
  logo_path = "logo.png"
  if os.path.exists(logo_path):
    logo_img = Image(logo_path, width=1.1 * inch, height=1.1 * inch)
    header_text = Paragraph(
        "<b>INJUL - Consolidado Maestro de CCTV</b><br/><font"
        " size=9 color='#555555'>Investigadores Jurídicos EIRL<br/>Fecha:"
        f" {datetime.now().strftime('%d/%m/%Y')} | Rango: 07:00 AM - 07:00"
        " PM</font>",
        title_style,
    )
    t_header = Table([[logo_img, header_text]], colWidths=[85, 455])
    t_header.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ])
    )
    story.append(t_header)
  else:
    story.append(Paragraph("INJUL - Consolidado Maestro de CCTV", title_style))

  story.append(Spacer(1, 15))

  kpi_data = [
      ["Cámaras Totales", "Operativas", "No Visibles", "Por Verificar", "Disponibilidad"],
      ["223", "121", "97", "5", "54.3%"],
  ]
  t_kpi = Table(kpi_data, colWidths=[100, 100, 100, 100, 140])
  t_kpi.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F497D")),
          ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
          ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F2F2F2")),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
      ])
  )
  story.append(t_kpi)
  story.append(PageBreak())

  # --- PÁGINA 2: HALLAZGOS ---
  story.append(Paragraph("<b>Hallazgos del Día y Resumen de Tickets</b>", title_style))
  story.append(Spacer(1, 10))
  story.append(
      Paragraph(
          "1. Monitoreo activo de la red de sucursales en EZStation.<br/>2."
          " Seguimiento de incidencias técnicas detalladas por cámara.",
          normal_style,
      )
  )
  story.append(PageBreak())

  # --- SUCURSALES Y SUS CÁMARAS EN EL PDF ---
  for idx, (suc, camaras) in enumerate(CATALOGO_CAMARAS.items(), start=1):
    story.append(Paragraph(f"<b>Sucursal {idx} de 9: {suc}</b>", title_style))
    story.append(Spacer(1, 10))

    branch_rows = [
        ["#", "Ubicación / Nombre en EZStation", "En vivo", "Grabación", "Estado", "Observaciones"]
    ]
    suc_data = datos_detalles.get(suc, {})

    for c_idx, cam_nombre in enumerate(camaras, start=1):
      info_cam = suc_data.get(cam_nombre, {"estado": "Operativa", "observacion": ""})
      estado_txt = info_cam["estado"]
      obs_txt = info_cam["observacion"]
      branch_rows.append(
          [str(c_idx), cam_nombre, "Sí", "Normal", estado_txt, obs_txt]
      )

    t_branch = Table(branch_rows, colWidths=[25, 185, 50, 70, 70, 140])
    t_branch.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#333333")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ])
    )
    story.append(t_branch)
    story.append(PageBreak())

  # --- PÁGINA FINAL: TICKETS ---
  story.append(Paragraph("<b>Seguimiento de Tickets con el Proveedor</b>", title_style))
  story.append(Spacer(1, 10))

  tickets_tabla_data = [["Ticket", "Fecha", "Sucursal / Área", "Estado", "Días", "Detalle"]]
  if lista_tickets_db:
    for t in lista_tickets_db:
      tickets_tabla_data.append([str(t[0]), str(t[1]), str(t[2]), str(t[3]), str(t[4]), str(t[5])])
  else:
    tickets_tabla_data.append(["N/A", "N/A", "Sin tickets registrados", "N/A", "0", "N/A"])

  t_tickets = Table(tickets_tabla_data, colWidths=[60, 75, 120, 80, 45, 150])
  t_tickets.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F497D")),
          ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
          ("FONTSIZE", (0, 0), (-1, -1), 8),
      ])
  )
  story.append(t_tickets)
  story.append(PageBreak())

  doc.build(story, canvasmaker=NumberedCanvas)
  return send_file(pdf_path, as_attachment=True)


@app.route("/descargar-excel")
def descargar_excel():
  excel_path = "Reporte_CCTV_INJUL.xlsx"
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Resumen"
  ws["A1"] = "INJUL - Consolidado de CCTV"
  ws["A1"].font = Font(name="Helvetica", size=14, bold=True, color="1F497D")
  wb.save(excel_path)
  return send_file(excel_path, as_attachment=True)


if __name__ == "__main__":
  inicializar_base_de_datos()
  port = int(os.environ.get("PORT", 8080))
  app.run(host="0.0.0.0", port=port)
