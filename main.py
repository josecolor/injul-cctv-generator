import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import psycopg2
from flask import Flask, redirect, render_template_string, request, send_file, url_for
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

# --- HTML DE LA INTERFAZ WEB CON FORMULARIO Y REGISTROS ---
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
        
        <!-- Formulario para ingresar datos -->
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

        <!-- Botones de Descarga -->
        <div class="card">
            <h3>Generación de Reportes Oficiales</h3>
            <a href="/descargar-pdf" class="btn">📥 Descargar PDF (13 Págs)</a>
            <a href="/descargar-excel" class="btn btn-excel">📊 Descargar Excel Matriz</a>
        </div>

        <!-- Historial Registrado -->
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

  doc.build(story)
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
