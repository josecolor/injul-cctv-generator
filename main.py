import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import psycopg2
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

# --- HTML DE LA INTERFAZ WEB ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>INJUL - Consolidado CCTV</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 15px; }
        .container { max-width: 900px; margin: 0 auto; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        h1 { color: #1F497D; font-size: 20px; border-bottom: 2px solid #1F497D; padding-bottom: 8px; }
        h3 { color: #1F497D; margin-top: 5px; }
        .card { background: #f8f9fa; padding: 15px; margin-bottom: 15px; border-radius: 6px; border-left: 4px solid #1F497D; }
        .btn { display: inline-block; background: #1F497D; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; font-weight: bold; margin-top: 5px; margin-right: 5px; border: none; cursor: pointer; font-size: 14px; }
        .btn-excel { background: #217346; }
        .btn:hover { opacity: 0.9; }
        .form-group { margin-bottom: 10px; }
        label { display: block; font-weight: bold; font-size: 13px; margin-bottom: 3px; color: #444; }
        input, select, textarea { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; overflow-x: auto; display: block; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; font-size: 12px; }
        th { background-color: #1F497D; color: white; }
    </style>
</head>
<body>
    <div class="container">
        <h1>INJUL - Sistema de Monitoreo CCTV</h1>
        
        <div class="card">
            <h3>Registrar Estado de Sucursal</h3>
            <form action="/guardar" method="POST">
                <div class="form-group">
                    <label>Sucursal:</label>
                    <select name="sucursal" required>
                        <option value="Caribbean Agregados">Caribbean Agregados</option>
                        <option value="Marmotech Blue Mall">Marmotech Blue Mall</option>
                        <option value="Marmotech DCM">Marmotech DCM</option>
                        <option value="Marmotech Outlet RP">Marmotech Outlet RP</option>
                        <option value="Marmotech Punta Cana">Marmotech Punta Cana</option>
                        <option value="Sucursal San Cristóbal">Sucursal San Cristóbal</option>
                        <option value="Marmotech Santiago 1">Marmotech Santiago 1</option>
                        <option value="Marmotech Santiago 2">Marmotech Santiago 2</option>
                        <option value="Santiago 2 - Sección Muebles">Santiago 2 - Sección Muebles</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Cámaras Operativas:</label>
                    <input type="number" name="operativas" required placeholder="Ej: 15">
                </div>
                <div class="form-group">
                    <label>Cámaras No Visibles / Sin Grabación:</label>
                    <input type="number" name="no_visibles" required placeholder="Ej: 2">
                </div>
                <div class="form-group">
                    <label>Observaciones / Novedades:</label>
                    <textarea name="observaciones" rows="2" placeholder="Detalles de cámaras o cableado..."></textarea>
                </div>
                <button type="submit" class="btn">💾 Guardar en Base de Datos</button>
            </form>
        </div>

        <div class="card">
            <h3>Generación de Reportes Oficiales</h3>
            <a href="/descargar-pdf" class="btn">📥 Descargar PDF Oficial (13 Págs)</a>
            <a href="/descargar-excel" class="btn btn-excel">📊 Descargar Excel Matriz</a>
        </div>

        <div class="card">
            <h3>Historial Reciente en PostgreSQL</h3>
            <table>
                <tr>
                    <th>Fecha</th>
                    <th>Sucursal</th>
                    <th>Operativas</th>
                    <th>No Visibles</th>
                    <th>Observaciones</th>
                </tr>
                {% for row in registros %}
                <tr>
                    <td>{{ row[1] }}</td>
                    <td>{{ row[2] }}</td>
                    <td>{{ row[3] }}</td>
                    <td>{{ row[4] }}</td>
                    <td>{{ row[5] }}</td>
                </tr>
                {% else %}
                <tr>
                    <td colspan="5">No hay registros guardados todavía.</td>
                </tr>
                {% endfor %}
            </table>
        </div>
    </div>
</body>
</html>
"""


# --- CLASE PARA MARCA DE AGUA Y PIE DE PÁGINA ---
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
    # Marca de agua CONFIDENCIAL
    self.saveState()
    self.setFont("Helvetica-Bold", 60)
    self.setFillColor(colors.HexColor("#F2DEDE"), alpha=0.3)
    self.translate(300, 400)
    self.rotate(45)
    self.drawCentredString(0, 0, "CONFIDENCIAL")
    self.restoreState()

    # Pie de página
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
            CREATE TABLE IF NOT EXISTS reportes_cctv (
                id SERIAL PRIMARY KEY,
                fecha DATE NOT NULL,
                sucursal VARCHAR(100) NOT NULL,
                operativas INT,
                no_visibles INT,
                observaciones TEXT,
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
  registros = []
  database_url = os.environ.get("DATABASE_URL")
  if database_url:
    try:
      conn = psycopg2.connect(database_url)
      cur = conn.cursor()
      cur.execute(
          "SELECT id, fecha, sucursal, operativas, no_visibles, observaciones"
          " FROM reportes_cctv ORDER BY id DESC LIMIT 15;"
      )
      registros = cur.fetchall()
      cur.close()
      conn.close()
    except Exception as e:
      print(f"Error leyendo DB: {e}")

  return render_template_string(HTML_TEMPLATE, registros=registros)


@app.route("/guardar", methods=["POST"])
def guardar():
  sucursal = request.form.get("sucursal")
  operativas = request.form.get("operativas")
  no_visibles = request.form.get("no_visibles")
  observaciones = request.form.get("observaciones")
  fecha_hoy = datetime.now().strftime("%Y-%m-%d")

  database_url = os.environ.get("DATABASE_URL")
  if database_url:
    try:
      conn = psycopg2.connect(database_url)
      cur = conn.cursor()
      cur.execute(
          """
                INSERT INTO reportes_cctv (fecha, sucursal, operativas, no_visibles, observaciones)
                VALUES (%s, %s, %s, %s, %s)
            """,
          (
              fecha_hoy,
              sucursal,
              int(operativas),
              int(no_visibles),
              observaciones,
          ),
      )
      conn.commit()
      cur.close()
      conn.close()
    except Exception as e:
      print(f"Error guardando en DB: {e}")

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

  # --- OBTENER DATOS DE LA BASE DE DATOS ---
  datos_sucursales = {}
  database_url = os.environ.get("DATABASE_URL")
  if database_url:
    try:
      conn = psycopg2.connect(database_url)
      cur = conn.cursor()
      # Traer el registro más reciente por sucursal
      cur.execute("""
                SELECT DISTINCT ON (sucursal) sucursal, operativas, no_visibles, observaciones, fecha
                FROM reportes_cctv
                ORDER BY sucursal, id DESC;
            """)
      rows = cur.fetchall()
      for r in rows:
        datos_sucursales[r[0]] = {
            "operativas": r[1],
            "no_visibles": r[2],
            "observaciones": r[3] or "",
            "fecha": str(r[4]),
        }
      cur.close()
      conn.close()
    except Exception as e:
      print(f"Error consultando DB para PDF: {e}")

  # --- PÁGINA 1: PORTADA Y RESUMEN DE KPIS ---
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

  # --- PÁGINA 2: HALLAZGOS Y SEGUIMIENTO ---
  story.append(Paragraph("<b>Hallazgos del Día y Resumen de Tickets</b>", title_style))
  story.append(Spacer(1, 10))
  hallazgos_texto = Paragraph(
      "1. Disponibilidad global de 54.3%: 121 de 223 cámaras operativas en"
      " EZStation.<br/>2. Hay 9 tickets en proceso con el proveedor; el más"
      " antiguo lleva 487 días abierto.<br/>3. Se recomienda revisión de"
      " cableado y conectividad en las sucursales con nivel crítico.",
      normal_style,
  )
  story.append(hallazgos_texto)
  story.append(PageBreak())

  # --- PÁGINAS 3 A 11: DESGLOSE DE LAS 9 SUCURSALES ---
  lista_sucursales = [
      "Caribbean Agregados",
      "Marmotech Blue Mall",
      "Marmotech DCM",
      "Marmotech Outlet RP",
      "Marmotech Punta Cana",
      "Sucursal San Cristóbal",
      "Marmotech Santiago 1",
      "Marmotech Santiago 2",
      "Santiago 2 - Sección Muebles",
  ]

  for idx, suc in enumerate(lista_sucursales, start=1):
    info = datos_sucursales.get(
        suc, {"operativas": 0, "no_visibles": 0, "observaciones": "Sin registro en base de datos"}
    )
    story.append(Paragraph(f"<b>Sucursal {idx} de 9: {suc}</b>", title_style))
    story.append(Spacer(1, 10))
    story.append(
        Paragraph(
            f"<b>Operativas:</b> {info['operativas']} | <b>No Visibles:</b>"
            f" {info['no_visibles']}<br/><b>Observaciones:</b>"
            f" {info['observaciones']}",
            normal_style,
        )
    )
    story.append(Spacer(1, 10))

    branch_data = [
        ["#", "Ubicación / Nombre en EZStation", "En vivo", "Grabación", "Estado", "Observaciones"],
        ["1", "Cámara Principal", "Sí", "Normal", "Operativa", info["observaciones"]],
    ]
    t_branch = Table(branch_data, colWidths=[30, 180, 60, 80, 80, 110])
    t_branch.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#333333")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
        ])
    )
    story.append(t_branch)
    story.append(PageBreak())

  # --- PÁGINA 12 Y 13: SEGUIMIENTO DE TICKETS CON EL PROVEEDOR ---
  story.append(Paragraph("<b>Seguimiento de Tickets con el Proveedor</b>", title_style))
  story.append(Spacer(1, 10))
  tickets_data = [
      ["Ticket", "Fecha", "Sucursal / Área", "Estado", "Días Abiertos"],
      ["9921", "20/06/2025", "San Cristóbal", "En proceso", "487"],
      ["9924", "20/06/2025", "Marmotech DCM", "En proceso", "457"],
  ]
  t_tickets = Table(tickets_data, colWidths=[70, 80, 130, 100, 90])
  t_tickets.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F497D")),
          ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
          ("FONTSIZE", (0, 0), (-1, -1), 9),
      ])
  )
  story.append(t_tickets)
  story.append(PageBreak())

  # Última página de cierre para completar el bloque
  story.append(Paragraph("<b>Cierre del Consolidado de CCTV</b>", title_style))
  story.append(Spacer(1, 15))
  story.append(
      Paragraph(
          "Reporte generado automáticamente a través del sistema MXL Digital"
          " Services / INJUL.",
          normal_style,
      )
  )

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
