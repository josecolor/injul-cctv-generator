import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


class NumberedCanvas(canvas.Canvas):
  """Canvas personalizado para agregar pie de página, número de página dinámico

  y marca de agua 'CONFIDENCIAL' en diagonal en todas las páginas.
  """

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

    # 1. Marca de agua CONFIDENCIAL en diagonal
    self.setFont("Helvetica-Bold", 60)
    self.setFillColor(colors.HexColor("#F2DEDE"), alpha=0.4)  # Rojo muy suave
    self.translate(300, 400)
    self.rotate(45)
    self.drawCentredString(0, 0, "CONFIDENCIAL")
    self.restoreState()

    self.saveState()
    self.setFont("Helvetica", 9)
    self.setFillColor(colors.HexColor("#555555"))

    # 2. Pie de página oficial INJUL
    footer_text_left = (
        "INJUL Consolidado Maestro de CCTV                         Documento"
        " confidencial"
    )
    footer_text_right = f"Página {self._pageNumber} de {total_pages}"

    self.drawString(36, 25, footer_text_left)
    self.drawRightString(letter[0] - 36, 25, footer_text_right)

    # Línea sutil sobre el pie de página
    self.setStrokeColor(colors.HexColor("#CCCCCC"))
    self.setLineWidth(0.5)
    self.line(36, 38, letter[0] - 36, 38)

    self.restoreState()


def generar_reporte_excel(output_excel_path):
  """Genera el libro Excel matriz con las 11 pestañas sincronizadas."""
  wb = openpyxl.Workbook()
  # Eliminar hoja por defecto
  default_sheet = wb.active
  wb.remove(default_sheet)

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
    ws.views.sheetView[0].showGridLines = True

    # Cabecera de prueba simulada
    ws["A1"] = f"INJUL - Consolidado de CCTV [{nombre}]"
    ws["A1"].font = Font(name="Helvetica", size=14, bold=True, color="1F497D")

    ws["A3"] = "#"
    ws["B3"] = "Ubicación / Cámara"
    ws["C3"] = "En Vivo"
    ws["D3"] = "Grabación"
    ws["E3"] = "Estado"
    ws["F3"] = "Observaciones"

    header_fill = PatternFill(
        start_color="1F497D", end_color="1F497D", fill_type="solid"
    )
    header_font = Font(name="Helvetica", size=10, bold=True, color="FFFFFF")

    for col in ["A", "B", "C", "D", "E", "F"]:
      cell = ws[f"{col}3"]
      cell.fill = header_fill
      cell.font = header_font
      cell.alignment = Alignment(horizontal="center", vertical="center")

  wb.save(output_excel_path)
  print(f"[+] Excel matriz guardado con éxito en: {output_excel_path}")


def generar_reporte_pdf(output_pdf_path):
  """Genera el PDF corporativo de 13 páginas estructurado exactamente al estándar de

  INJUL.
  """
  doc = SimpleDocTemplate(
      output_pdf_path,
      pagesize=letter,
      leftMargin=36,
      rightMargin=36,
      topMargin=36,
      bottomMargin=50,
  )

  story = []
  styles = getSampleStyleSheet() if "getSampleStyleSheet" in globals() else None
  # Definición simple de estilos por si acaso
  from reportlab.lib.styles import ParagraphStyle

  title_style = ParagraphStyle(
      "TitleStyle", fontName="Helvetica-Bold", fontSize=16, textColor=colors.HexColor("#1F497D")
  )
  normal_style = ParagraphStyle(
      "NormalStyle", fontName="Helvetica", fontSize=10, textColor=colors.HexColor("#333333")
  )

  # --- PÁGINA 1: PANEL DE CONTROL Y RESUMEN ---
  story.append(Paragraph("INJUL - Consolidado Maestro de CCTV", title_style))
  story.append(Spacer(1, 10))
  story.append(
      Paragraph(
          "<b>Fecha del Reporte:</b> 23/10/2026 | <b>Rango:</b> 07:00 AM - 07:00"
          " PM",
          normal_style,
      )
  )
  story.append(Spacer(1, 15))

  # Tabla resumen KPIs globales
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
          ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
          ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F2F2F2")),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
      ])
  )
  story.append(t_kpi)
  story.append(PageBreak())

  # --- PÁGINAS 2 A 10: SUCURSALES (Ejemplo estructurado dinámico) ---
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
    story.append(
        Paragraph(
            f"<b>Sucursal {idx} de 9: {suc}</b>", title_style
        )
    )
    story.append(Spacer(1, 10))

    # Tabla de cámaras de ejemplo para la sucursal
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

  # --- PÁGINA 13: SEGUIMIENTO DE TICKETS ---
  story.append(Paragraph("Seguimiento Detallado de Tickets con el Proveedor", title_style))
  story.append(Spacer(1, 10))
  ticket_data = [
      ["Ticket", "Fecha", "Sucursal / Área", "Estado", "Días Abiertos"],
      ["9921", "20/06/2025", "San Cristóbal", "En proceso", "487"],
      ["9924", "20/06/2025", "Marmotech DCM", "En proceso", "457"],
  ]
  t_tickets = Table(ticket_data, colWidths=[70, 90, 150, 100, 130])
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

  # Construir documento usando el canvas numerado con la marca de agua
  doc.build(story, canvasmaker=NumberedCanvas)
  print(f"[+] PDF corporativo guardado con éxito en: {output_pdf_path}")


if __name__ == "__main__":
  fecha_hoy = datetime.now().strftime("%Y-%m-%d")
  pdf_nombre = f"Reporte_CCTV_INJUL_{fecha_hoy}.pdf"
  excel_nombre = f"Reporte_CCTV_INJUL_{fecha_hoy}.xlsx"

  print("[*] Iniciando generación de archivos...")
  generar_reporte_excel(excel_nombre)
  generar_reporte_pdf(pdf_nombre)
  print("[*] ¡Proceso completado con éxito!")
