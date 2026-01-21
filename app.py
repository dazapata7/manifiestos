import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
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
    FECHA_MANIFIESTO = datetime.now().strftime('%d/%m/%Y')
    st.info(f"📅 Fecha: **{FECHA_MANIFIESTO}**")
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
                # Crear PDF
                buffer = BytesIO()
                doc = SimpleDocTemplate(
                    buffer,
                    pagesize=landscape(letter),
                    rightMargin=20,
                    leftMargin=20,
                    topMargin=40,
                    bottomMargin=30
                )
                
                elements = []
                styles = getSampleStyleSheet()
                
                # Estilos simples
                title_style = ParagraphStyle(
                    'Title',
                    parent=styles['Heading1'],
                    fontSize=14,
                    alignment=TA_CENTER,
                    spaceAfter=6,
                    fontName='Helvetica-Bold'
                )
                
                subtitle_style = ParagraphStyle(
                    'Subtitle',
                    parent=styles['Normal'],
                    fontSize=9,
                    alignment=TA_CENTER,
                    spaceAfter=12,
                    fontName='Helvetica'
                )
                
                # Anchos de columnas optimizados
                col_widths = [
                    0.4 * inch,   # #
                    0.8 * inch,   # Guía
                    1.6 * inch,   # Cliente
                    1.0 * inch,   # Ciudad
                    1.0 * inch,   # Estado
                    1.8 * inch,   # Dirección
                    2.0 * inch    # Producto
                ]
                
                total_ordenes = len(df)
                ordenes_por_pagina = 18
                num_paginas = (total_ordenes + ordenes_por_pagina - 1) // ordenes_por_pagina
                
                # Procesar cada página
                for pagina in range(num_paginas):
                    inicio = pagina * ordenes_por_pagina
                    fin = min((pagina + 1) * ordenes_por_pagina, total_ordenes)
                    
                    if pagina > 0:
                        elements.append(PageBreak())
                    
                    # Encabezado
                    elements.append(Paragraph("MANIFIESTO DE ENTREGA", title_style))
                    
                    if num_paginas > 1:
                        elements.append(Paragraph(
                            f"Fecha: {FECHA_MANIFIESTO} | Total: {total_ordenes} órdenes | Página {pagina + 1} de {num_paginas}",
                            subtitle_style
                        ))
                    else:
                        elements.append(Paragraph(
                            f"Fecha: {FECHA_MANIFIESTO} | Total: {total_ordenes} órdenes",
                            subtitle_style
                        ))
                    
                    elements.append(Spacer(1, 0.2 * inch))
                    
                    # Preparar datos de la tabla
                    chunk = df.iloc[inicio:fin]
                    table_data = []
                    
                    # ENCABEZADOS
                    table_data.append(['#', 'Guía', 'Cliente', 'Ciudad', 'Estado', 'Dirección', 'Producto'])
                    
                    # DATOS (SIN PARAGRAPH, solo texto simple)
                    for idx, row in chunk.iterrows():
                        numero_orden = inicio + (idx - chunk.index[0]) + 1
                        
                        # Limitar textos pero permitir que se expandan verticalmente
                        guia = str(row['Guía de Envío']) if pd.notna(row['Guía de Envío']) else ''
                        cliente = str(row['Cliente'])[:30] if pd.notna(row['Cliente']) else ''
                        ciudad = str(row['Ciudad'])[:15] if pd.notna(row['Ciudad']) else ''
                        estado = str(row['Estado'])[:12] if pd.notna(row['Estado']) else ''
                        
                        # Dirección (sin truncar)
                        direccion = ''
                        if pd.notna(row['Calle']):
                            direccion = str(row['Calle'])
                        if pd.notna(row['Número']):
                            direccion += ' ' + str(row['Número'])
                        direccion = direccion.strip()
                        
                        # Producto (sin truncar)
                        producto = str(row['Productos']) if pd.notna(row['Productos']) else ''
                        
                        table_data.append([
                            str(numero_orden),
                            guia,
                            cliente,
                            ciudad,
                            estado,
                            direccion,
                            producto
                        ])
                    
                    # Crear tabla SIN rowHeights fijo (se expande automáticamente)
                    tabla = Table(table_data, colWidths=col_widths, repeatRows=1)
                    
                    # Estilos MINIMALISTAS pero efectivos
                    estilo = TableStyle([
                        # Encabezado
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 9),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                        ('TOPPADDING', (0, 0), (-1, 0), 8),
                        
                        # Datos
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                        ('ALIGN', (0, 1), (1, -1), 'CENTER'),
                        ('ALIGN', (2, 1), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        
                        # Bordes
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                        
                        # Filas alternadas
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
                    ])
                    
                    tabla.setStyle(estilo)
                    elements.append(tabla)
                
                # Página de firmas (solo una)
                elements.append(PageBreak())
                elements.append(Spacer(1, 2 * inch))
                
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
                
                firma_table = Table(firma_data, colWidths=[3 * inch, 0.5 * inch, 0.5 * inch, 3 * inch])
                firma_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 2), (-1, 2), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 11),
                ]))
                
                elements.append(firma_table)
                
                # Generar PDF
                doc.build(elements)
                buffer.seek(0)
                
                # Descargar
                st.success(f"✅ PDF generado: {num_paginas} páginas de datos + 1 página de firmas")
                
                st.download_button(
                    label="📥 Descargar PDF",
                    data=buffer.getvalue(),
                    file_name=nombre_pdf,
                    mime="application/pdf",
                    use_container_width=True
                )
                
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
