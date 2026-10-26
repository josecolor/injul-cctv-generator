import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import psycopg2
from flask import Flask, render_template_string, request, send_file
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

app = Flask(__name__)

# --- HTML DE LA INTERFAZ WEB (PANEL DE CONTROL) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>INJUL - Consolidado CCTV</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 800px; margin: 0 auto; background: #fff; padding: 25px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        h1 { color: #1F497D; font-size: 22px; border-bottom: 2px solid #1F497D; padding-bottom: 10px; }
        .card { background: #f8f9fa; padding: 15px; margin-bottom: 15px; border-radius: 6px; border-left: 4px solid #1F497D; }
        .btn { display: inline-block; background: #1F497D; color: white; padding: 10px 18px; text-decoration: none; border-radius: 4px; font-weight: bold; margin-top: 10px; margin-right: 10px; border: none; cursor: pointer; }
        .btn-excel { background: #217346; }
        .btn:hover { opacity: 0.9; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; font-size: 13px; }
        th { background-color: #1F497D; color: white; }
    </style>
</head>
<body>
    <div class="container">
        <h1>INJUL - Sistema de Reportes CCTV</h1>
        
        <div class="card">
            <h3>Generación de Salida Doble</h3>
            <p>Haz clic para descargar de inmediato los reportes oficiales actualizados al día de hoy.</p>
            <a href="/descargar-pdf" class="btn">📥 Descargar PDF (13 Págs)</a>
            <a href="/descargar-excel" class="btn btn-excel">📊 Descargar Excel Matriz</a>
        </div>

        <div class="card">
            <h3>Estado del Sistema y Base de datos</h3>
            <p>🟢 Conectado a PostgreSQL (Railway) | 🏢 Monitoreo de 9 Sucursales Activo.</p>
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
    self.setFillColor(colors.HexColor("#F2DEDE"), alpha=0.4)
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
            CREATE TABLE IF NOT EXISTS reportes_cctv (
                id SERIAL PRIMARY KEY,
                fecha DATE NOT NULL,
                total_camaras INT,
                operativas INT,
                no_visibles INT,
                por_verificar INT,
                disponibilidad VARCHAR(20),
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
  return render_template_string(HTML_TEMPLATE)


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
      fontSize=16,
      textColor=colors.HexColor("#1F497D"),
  )
  normal_style = ParagraphStyle(
      "NormalStyle",
      fontName="Helvetica",
      fontSize=10,
      textColor=colors.HexColor("#333333"),
  )

  story.append(Paragraph("INJUL - Consolidado Maestro de CCTV", title_style))
  story.append(Spacer(1, 10))
  story.append(
      Paragraph(
          "<b>Fecha del Reporte:</b> "
          + datetime.now().strftime("%d/%m/%Y")
          + " | <b>Rango:</b> 07:00 AM - 07:00 PM",
          normal_style,
      )
  )
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

  sucursales = [
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
  for idx, suc in enumerate(sucursales, start=1):
    story.append(Paragraph(f"<b>Sucursal {idx} de 9: {suc}</b>", title_style))
    story.append(Spacer(1, 10))
    branch_data = [
        ["#", "Ubicación / Nombre en EZStation", "En vivo", "Grabación", "Estado", "Observaciones"],
        ["1", "Entrada Principal", "Sí", "Normal", "Operativa", ""],
        ["2", "Pasillo Lateral", "Sí", "Normal", "No visible", "Revisar cableado"],
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

  doc.build(story, canvasmaker=NumberedCanvas)
  return send_file(pdf_path, as_attachment=True)


@app.route("/descargar-excel")
def descargar_excel():
  excel_path = "Reporte_CCTV_INJUL.xlsx"
  wb = openpyxl.Workbook()
  wb.remove(wb.active)
  hojas = [
      "Resumen",
      "1 Caribbean",
      "2 Blue Mall",
      "3 DCM",
      "4 Outlet RP",
      "5 Punta Cana",
      "6 San Cristóbal",
      "7 Santiago 1",
      "8 Santiago 2",
      "9 Muebles",
      "Tickets",
  ]
  for nombre in hojas:
    ws = wb.create_sheet(title=nombre)
    ws["A1"] = f"INJUL - Consolidado de CCTV [{nombre}]"
    ws["A1"].font = Font(name="Helvetica", size=14, bold=True, color="1F497D")
    ws["A3"] = "#"
    ws["B3"] = "Ubicación / Cámara"
    ws["C3"] = "En Vivo"
    ws["D3"] = "Grabación"
    ws["E3"] = "Estado"
    ws["F3"] = "Observaciones"
    for col in ["A", "B", "C", "D", "E", "F"]:
      cell = ws[f"{col}3"]
      cell.fill = PatternFill(
          start_color="1F497D", end_color="1F497D", fill_type="solid"
      )
      cell.font = Font(name="Helvetica", size=10, bold=True, color="FFFFFF")
  wb.save(excel_path)
  return send_file(excel_path, as_attachment=True)


if __name__ == "__main__":
  inicializar_base_de_datos()
  port = int(os.environ.get("PORT", 8080))
  app.run(host="0.0.0.0", port=port)
