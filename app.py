import streamlit as st
import pandas as pd
import base64
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="Generador de Manifiestos",
    page_icon="📦",
    layout="wide"
)

# Título
st.title("📦 Generador de Manifiestos de Entrega")

# Sidebar
with st.sidebar:
    st.header("⚙️ Configuración")
    fecha_option = st.radio("Fecha del manifiesto:", ["Fecha actual", "Especificar fecha"])
    
    if fecha_option == "Especificar fecha":
        fecha_manual = st.date_input("Selecciona fecha:", datetime.now())
        FECHA_MANIFIESTO = fecha_manual.strftime('%d/%m/%Y')
    else:
        FECHA_MANIFIESTO = datetime.now().strftime('%d/%m/%Y')
    
    nombre_pdf = st.text_input("Nombre del PDF:", f"Manifiesto_{FECHA_MANIFIESTO.replace('/', '_')}.pdf")

# Subir archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Verificar columnas
        columnas_requeridas = ['Guía de Envío', 'Cliente', 'Ciudad', 'Estado', 'Calle', 'Número', 'Productos']
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            st.error(f"❌ Columnas faltantes: {', '.join(columnas_faltantes)}")
            st.stop()
        
        st.success(f"✅ {len(df)} órdenes cargadas")
        
        if st.button("🔄 Generar PDF", type="primary"):
            with st.spinner("Generando PDF..."):
                buffer = BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), 
                                       rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
                
                elements = []
                styles = getSampleStyleSheet()
                
                # Estilos
                title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=14,
                    textColor=colors.black, spaceAfter=6, alignment=TA_CENTER, fontName='Helvetica-Bold')
                
                subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'], fontSize=9,
                    textColor=colors.grey, spaceAfter=10, alignment=TA_CENTER, fontName='Helvetica')
                
                # Anchos de columnas fijos
                col_widths = [0.4*inch, 0.8*inch, 1.6*inch, 1.0*inch, 1.0*inch, 1.8*inch, 2.0*inch]
                
                # Función para crear celdas con texto que se expande
                def crear_celda(texto, ancho, estilo='normal'):
                    if pd.isna(texto):
                        texto = ''
                    else:
                        texto = str(texto)
                    
                    # Crear Paragraph que permite múltiples líneas
                    if estilo == 'numero':
                        para = Paragraph(texto, ParagraphStyle('Cell', 
                            parent=styles['Normal'], fontSize=8, alignment=TA_CENTER,
                            fontName='Helvetica', leading=9))
                    else:
                        para = Paragraph(texto, ParagraphStyle('Cell', 
                            parent=styles['Normal'], fontSize=8, alignment=TA_LEFT,
                            fontName='Helvetica', leading=9, wordWrap='CJK'))
                    
                    return para
                
                # Procesar datos en chunks dinámicos
                total_ordenes = len(df)
                ordenes_por_pagina = 18  # Objetivo inicial
                elementos_por_pagina = []
                pagina_actual = []
                
                for idx, row in df.iterrows():
                    # Preparar datos de la fila
                    fila = [
                        crear_celda(idx + 1, col_widths[0], 'numero'),
                        crear_celda(row['Guía de Envío'], col_widths[1], 'numero'),
                        crear_celda(row['Cliente'], col_widths[2]),
                        crear_celda(row['Ciudad'], col_widths[3]),
                        crear_celda(row['Estado'], col_widths[4]),
                        crear_celda(f"{row['Calle']} {row['Número']}".strip(), col_widths[5]),
                        crear_celda(row['Productos'], col_widths[6])
                    ]
                    
                    pagina_actual.append(fila)
                    
                    # Si tenemos 18 órdenes o el texto es muy largo, crear nueva página
                    if len(pagina_actual) >= ordenes_por_pagina:
                        elementos_por_pagina.append(pagina_actual)
                        pagina_actual = []
                
                # Agregar la última página si tiene datos
                if pagina_actual:
                    elementos_por_pagina.append(pagina_actual)
                
                # Generar páginas
                num_paginas = len(elementos_por_pagina)
                
                for pagina_num, datos_pagina in enumerate(elementos_por_pagina):
                    if pagina_num > 0:
                        elements.append(PageBreak())
                    
                    # Encabezado
                    elements.append(Paragraph("MANIFIESTO DE ENTREGA", title_style))
                    elements.append(Paragraph(f"Fecha: {FECHA_MANIFIESTO} | Total: {total_ordenes} órdenes | Página {pagina_num+1} de {num_paginas}", subtitle_style))
                    elements.append(Spacer(1, 0.2*inch))
                    
                    # Crear tabla con encabezados
                    table_data = [[
                        Paragraph('#', ParagraphStyle('Header', fontSize=9, alignment=TA_CENTER, fontName='Helvetica-Bold')),
                        Paragraph('Guía', ParagraphStyle('Header', fontSize=9, alignment=TA_CENTER, fontName='Helvetica-Bold')),
                        Paragraph('Cliente', ParagraphStyle('Header', fontSize=9, alignment=TA_CENTER, fontName='Helvetica-Bold')),
                        Paragraph('Ciudad', ParagraphStyle('Header', fontSize=9, alignment=TA_CENTER, fontName='Helvetica-Bold')),
                        Paragraph('Estado', ParagraphStyle('Header', fontSize=9, alignment=TA_CENTER, fontName='Helvetica-Bold')),
                        Paragraph('Dirección', ParagraphStyle('Header', fontSize=9, alignment=TA_CENTER, fontName='Helvetica-Bold')),
                        Paragraph('Producto', ParagraphStyle('Header', fontSize=9, alignment=TA_CENTER, fontName='Helvetica-Bold'))
                    ]]
                    
                    table_data.extend(datos_pagina)
                    
                    # Crear tabla sin altura fija (se expande automáticamente)
                    tabla = Table(table_data, colWidths=col_widths, repeatRows=1)
                    
                    tabla.setStyle(TableStyle([
                        # Encabezado
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 9),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                        ('TOPPADDING', (0, 0), (-1, 0), 6),
                        
                        # Datos
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),
                        ('BOTTOMPADDING', (0, 1), (-1, -1), 3),
                        ('TOPPADDING', (0, 1), (-1, -1), 3),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        
                        # Bordes
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                        
                        # Filas alternadas
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f9f9f9')]),
                    ]))
                    
                    elements.append(tabla)
                
                # Página de firmas
                elements.append(PageBreak())
                elements.append(Spacer(1, 2*inch))
                
                firma_data = [
                    ['', '', '', ''],
                    ['_________________________', '', '', '_________________________'],
                    ['Entregado por', '', '', 'Recibido por'],
                    ['', '', '', ''],
                    ['Nombre:', '', '', 'Nombre:'],
                    ['', '', '', ''],
                    ['Fecha:', '', '', 'Fecha:'],
                    ['', '', '', ''],
                    ['Hora:', '', '', 'Hora:'],
                ]
                
                firma_table = Table(firma_data, colWidths=[3*inch, 0.5*inch, 0.5*inch, 3*inch])
                firma_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 2), (-1, 2), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 11),
                ]))
                
                elements.append(firma_table)
                elements.append(Spacer(1, 0.8*inch))
                
                nota = Paragraph("Documento generado automáticamente.", 
                    ParagraphStyle('Nota', fontSize=9, textColor=colors.grey, alignment=TA_CENTER))
                elements.append(nota)
                
                # Generar PDF
                doc.build(elements)
                buffer.seek(0)
                pdf_data = buffer.getvalue()
                
                # Mostrar resultado
                st.success(f"✅ PDF generado: {len(elementos_por_pagina)} páginas")
                
                st.download_button(
                    label="📥 Descargar PDF",
                    data=pdf_data,
                    file_name=nombre_pdf,
                    mime="application/pdf",
                    use_container_width=True
                )
                
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
