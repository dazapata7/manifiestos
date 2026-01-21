import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas as rl_canvas
from xml.sax.saxutils import escape


# -----------------------------
# Canvas con encabezado + "Página X de Y"
# -----------------------------
class HeaderCanvas(rl_canvas.Canvas):
    def __init__(self, *args, manifest_date="", total_orders=0, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        self.manifest_date = manifest_date
        self.total_orders = total_orders

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_pages = len(self._saved_page_states)
        for page_num, state in enumerate(self._saved_page_states, start=1):
            self.__dict__.update(state)
            self.draw_header(page_num, total_pages)
            super().showPage()
        super().save()

    def draw_header(self, page_num, total_pages):
        width, height = self._pagesize

        self.setFont("Helvetica-Bold", 14)
        self.drawCentredString(width / 2.0, height - 22, "MANIFIESTO DE ENTREGA")

        self.setFont("Helvetica", 9)
        subtitle = (
            f"Fecha: {self.manifest_date} | Total: {self.total_orders} órdenes | "
            f"Página {page_num} de {total_pages}"
        )
        self.drawCentredString(width / 2.0, height - 36, subtitle)


def as_para(text: str, style: ParagraphStyle) -> Paragraph:
    """
    Convierte texto a Paragraph para:
    - envolver en la celda (sin truncar ni "...")
    - respetar saltos de línea
    """
    if text is None:
        text = ""
    text = str(text)
    text = escape(text).replace("\n", "<br/>")
    return Paragraph(text, style)


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Generador de Manifiestos", page_icon="📦", layout="wide")
st.title("📦 Generador de Manifiestos de Entrega")

with st.sidebar:
    st.header("⚙️ Configuración")
    FECHA_MANIFIESTO = datetime.now().strftime("%d/%m/%Y")
    st.info(f"📅 Fecha: **{FECHA_MANIFIESTO}**")
    nombre_pdf = st.text_input(
        "Nombre del PDF:", f"Manifiesto_{FECHA_MANIFIESTO.replace('/', '_')}.pdf"
    )

uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        # Ajusta si tus columnas se llaman distinto
        columnas_requeridas = [
            "Guía de Envío",
            "Cliente",
            "Ciudad",
            "Estado",
            "Calle",
            "Número",
            "Productos",
        ]
        faltantes = [c for c in columnas_requeridas if c not in df.columns]
        if faltantes:
            st.error(f"❌ Columnas faltantes: {', '.join(faltantes)}")
            st.stop()

        total_ordenes = len(df)
        st.success(f"✅ {total_ordenes} órdenes cargadas")

        if st.button("🔄 Generar PDF", type="primary"):
            with st.spinner("Generando PDF..."):
                buffer = BytesIO()

                # Carta horizontal + márgenes para impresión
                doc = SimpleDocTemplate(
                    buffer,
                    pagesize=landscape(letter),
                    leftMargin=40,
                    rightMargin=40,
                    topMargin=60,      # deja espacio para header del canvas
                    bottomMargin=30,
                )

                styles = getSampleStyleSheet()

                # Estilo de celdas (wrap)
                cell_style = ParagraphStyle(
                    "Cell",
                    parent=styles["Normal"],
                    fontName="Helvetica",
                    fontSize=8,
                    leading=9,
                    spaceBefore=0,
                    spaceAfter=0,
                    wordWrap="CJK",
                )

                # Header blanco + más legible
                header_style = ParagraphStyle(
                    "HeaderCell",
                    parent=styles["Normal"],
                    fontName="Helvetica-Bold",
                    fontSize=10,
                    leading=11,
                    alignment=TA_CENTER,
                    textColor=colors.white,
                )

                # Anchos responsivos según el ancho útil del documento (con márgenes)
                available_w = doc.width
                ratios = [0.04, 0.10, 0.17, 0.11, 0.11, 0.25, 0.22]  # suman 1.0
                col_widths = [available_w * r for r in ratios]

                # Tabla completa (se parte sola entre páginas)
                table_data = [
                    [
                        as_para("#", header_style),
                        as_para("Guía", header_style),
                        as_para("Cliente", header_style),
                        as_para("Ciudad", header_style),
                        as_para("Estado", header_style),
                        as_para("Dirección", header_style),
                        as_para("Producto", header_style),
                    ]
                ]

                for i, row in df.iterrows():
                    guia = row.get("Guía de Envío", "")
                    cliente = row.get("Cliente", "")
                    ciudad = row.get("Ciudad", "")
                    estado = row.get("Estado", "")

                    calle = row.get("Calle", "")
                    numero = row.get("Número", "")
                    direccion = f"{calle} {numero}".strip()

                    producto = row.get("Productos", "")

                    table_data.append(
                        [
                            as_para(str(i + 1), cell_style),
                            as_para(guia, cell_style),
                            as_para(cliente, cell_style),
                            as_para(ciudad, cell_style),
                            as_para(estado, cell_style),
                            as_para(direccion, cell_style),
                            as_para(producto, cell_style),
                        ]
                    )

                tabla = Table(
                    table_data,
                    colWidths=col_widths,
                    repeatRows=1,
                    splitByRow=1,
                )

                tabla.setStyle(
                    TableStyle(
                        [
                            # Header
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, 0), 8),
                            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),

                            # Body
                            ("VALIGN", (0, 1), (-1, -1), "TOP"),
                            ("TOPPADDING", (0, 1), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 1), (-1, -1), 4),

                            # Grid
                            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),

                            # Alternating rows
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
                        ]
                    )
                )

                elements = []
                elements.append(Spacer(1, 0.10 * inch))
                elements.append(tabla)

                # Página de firmas
                elements.append(PageBreak())
                elements.append(Spacer(1, 2.0 * inch))

                firma_data = [
                    ["", "", "", ""],
                    ["_________________________", "", "", "_________________________"],
                    ["Entregado por", "", "", "Recibido por"],
                    ["Nombre:", "", "", "Nombre:"],
                    ["Fecha:", "", "", "Fecha:"],
                    ["Hora:", "", "", "Hora:"],
                ]

                firma_table = Table(
                    firma_data,
                    colWidths=[3 * inch, 0.5 * inch, 0.5 * inch, 3 * inch],
                )
                firma_table.setStyle(
                    TableStyle(
                        [
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("FONTNAME", (0, 2), (-1, 2), "Helvetica-Bold"),
                            ("FONTSIZE", (0, 0), (-1, -1), 11),
                        ]
                    )
                )

                elements.append(firma_table)
                elements.append(Spacer(1, 0.3 * inch))

                note_style = ParagraphStyle(
                    "Note",
                    parent=styles["Normal"],
                    fontSize=9,
                    alignment=TA_CENTER,
                )
                elements.append(
                    Paragraph(
                        "Este documento es un manifiesto de entrega generado automáticamente. "
                        "Para cualquier aclaración, contactar con el área de logística.",
                        note_style,
                    )
                )

                def canvasmaker(*args, **kwargs):
                    return HeaderCanvas(
                        *args,
                        manifest_date=FECHA_MANIFIESTO,
                        total_orders=total_ordenes,
                        **kwargs,
                    )

                doc.build(elements, canvasmaker=canvasmaker)

                buffer.seek(0)
                st.success("✅ PDF generado correctamente")
                st.download_button(
                    label="📥 Descargar PDF",
                    data=buffer.getvalue(),
                    file_name=nombre_pdf,
                    mime="application/pdf",
                    use_container_width=True,
                )

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
