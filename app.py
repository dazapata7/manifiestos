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
from reportlab.platypus.flowables import KeepTogether

# Configuración de la página
st.set_page_config(
    page_title="Generador de Manifiestos",
    page_icon="📦",
    layout="wide"
)

# Título
st.title("📦 Generador de Manifiestos de Entrega")
st.markdown("Sube tu archivo Excel y descarga el PDF automáticamente")

# Sidebar con configuraciones
with st.sidebar:
    st.header("⚙️ Configuración")
    
    fecha_option = st.radio("Fecha del manifiesto:", 
                           ["Fecha actual", "Especificar fecha"])
    
    if fecha_option == "Especificar fecha":
        fecha_manual = st.date_input("Selecciona fecha:", datetime.now())
        FECHA_MANIFIESTO = fecha_manual.strftime('%d/%m/%Y')
    else:
        FECHA_MANIFIESTO = datetime.now().strftime('%d/%m/%Y')
    
    st.info(f"📅 Fecha: **{FECHA_MANIFIESTO}**")
    
    nombre_pdf = st.text_input("Nombre del PDF:", 
                              f"Manifiesto_{FECHA_MANIFIESTO.replace('/', '_')}.pdf")
    
    st.markdown("---")
    st.markdown("### 📋 Columnas requeridas:")
    st.markdown("""
    - Guía de Envío
    - Cliente
    - Ciudad
    - Estado
    - Calle
    - Número
    - Productos
    """)

# Área principal - Subir archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=['xlsx', 'xls'], 
                                 help="Asegúrate de que tenga las columnas requeridas")

if uploaded_file is not None:
    try:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file)
        
        # Verificar columnas requeridas
        columnas_requeridas = ['Guía de Envío', 'Cliente', 'Ciudad', 'Estado', 
                              'Calle', 'Número', 'Productos']
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            st.error(f"❌ Columnas faltantes: {', '.join(columnas_faltantes)}")
            st.stop()
        
        # Mostrar vista previa
        with st.expander("👁️ Vista previa de datos", expanded=True):
            st.dataframe(df[columnas_requeridas].head(), use_container_width=True)
        
        st.success(f"✅ Archivo cargado - {len(df)} órdenes encontradas")
        
        # Botón para generar PDF
        if st.button("🔄 Generar PDF", type="primary", use_container_width=True):
            with st.spinner("Generando PDF..."):
                # Crear PDF en memoria
                buffer = BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), 
                                       rightMargin=20, leftMargin=20,
                                       topMargin=30, bottomMargin=30)
                
                elements = []
                styles = getSampleStyleSheet()
                
                # Estilos
                title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], 
                    fontSize=16, textColor=colors.HexColor('#1a1a1a'), 
                    spaceAfter=8, alignment=TA_CENTER, fontName='Helvetica-Bold')
                
                subtitle_style = ParagraphStyle('CustomSubtitle', parent=styles['Normal'], 
                    fontSize=10, textColor=colors.HexColor('#666666'), 
                    spaceAfter=12, alignment=TA_CENTER, fontName='Helvetica')
                
                # ANCHOS DE COLUMNAS OPTIMIZADOS (corregidos)
                col_widths = [
                    0.4*inch,    # # (más estrecho)
                    0.9*inch,    # Guía (más estrecho)
                    1.6*inch,    # Cliente (más ancho)
                    1.0*inch,    # Ciudad
                    1.0*inch,    # Estado
                    1.8*inch,    # Dirección
                    2.2*inch     # Producto (más ancho para múltiples líneas)
                ]
                
                # Calcular páginas
                total_ordenes = len(df)
                ordenes_por_pagina = 18
                num_paginas = (total_ordenes + ordenes_por_pagina - 1) // ordenes_por_pagina
                
                # Generar páginas
                for i in range(num_paginas):
                    start = i * ordenes_por_pagina
                    end = min((i + 1) * ordenes_por_pagina, total_ordenes)
                    
                    if i > 0:
                        elements.append(PageBreak())
                    
                    # Encabezado
                    elements.append(Paragraph("MANIFIESTO DE ENTREGA", title_style))
                    if num_paginas > 1:
                        elements.append(Paragraph(f"Fecha: {FECHA_MANIFIESTO} | Total: {total_ordenes} órdenes | Página {i+1} de {num_paginas}", subtitle_style))
                    else:
                        elements.append(Paragraph(f"Fecha: {FECHA_MANIFIESTO} | Total: {total_ordenes} órdenes", subtitle_style))
                    
                    # Datos de la tabla
                    chunk = df.iloc[start:end]
                    table_data = [['#', 'Guía', 'Cliente', 'Ciudad', 'Estado', 'Dirección', 'Producto']]
                    
                    for idx, row in chunk.iterrows():
                        guia = str(row['Guía de Envío']) if pd.notna(row['Guía de Envío']) else 'N/A'
                        cliente = str(row['Cliente'])[:25] if pd.notna(row['Cliente']) else 'N/A'
                        ciudad = str(row['Ciudad'])[:12] if pd.notna(row['Ciudad']) else 'N/A'
                        estado = str(row['Estado'])[:10] if pd.notna(row['Estado']) else 'N/A'
                        
                        # Producto con ajuste de líneas automático
                        if pd.notna(row['Productos']):
                            producto_texto = str(row['Productos'])
                            # Crear Paragraph que permite múltiples líneas
                            producto_para = Paragraph(producto_texto, 
                                ParagraphStyle('Producto', parent=styles['Normal'], 
                                fontSize=7, alignment=TA_LEFT, fontName='Helvetica',
                                wordWrap='CJK'))  # Permite wrap de texto
                        else:
                            producto_para = Paragraph('N/A', 
                                ParagraphStyle('Producto', parent=styles['Normal'], 
                                fontSize=7, alignment=TA_LEFT, fontName='Helvetica'))
                        
                        # Dirección
                        direccion_parts = []
                        if pd.notna(row['Calle']):
                            direccion_parts.append(str(row['Calle']))
                        if pd.notna(row['Número']):
                            direccion_parts.append(str(row['Número']))
                        direccion = ' '.join(direccion_parts)[:30] if direccion_parts else 'N/A'
                        
                        # NUMERACIÓN CORREGIDA (usa el índice global)
                        numero_orden = start + (idx - chunk.index[0]) + 1
                        
                        table_data.append([
                            str(numero_orden),  # Numeración correcta
                            guia,
                            cliente,
                            ciudad,
                            estado,
                            direccion,
                            producto_para  # Usa Paragraph para múltiples líneas
                        ])
                    
                    # Crear tabla con estilos optimizados
                    guias_table = Table(table_data, colWidths=col_widths, repeatRows=1)
                    guias_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 9),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                        ('TOPPADDING', (0, 0), (-1, 0), 8),
                        
                        # Filas de datos
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                        ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # Columna #
                        ('ALIGN', (1, 1), (1, -1), 'CENTER'),  # Columna Guía
                        ('ALIGN', (2, 1), (5, -1), 'LEFT'),    # Cliente, Ciudad, Estado, Dirección
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (5, -1), 8),      # Tamaño para todas excepto Producto
                        ('FONTSIZE', (6, 1), (6, -1), 7),      # Producto más pequeño
                        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
                        ('TOPPADDING', (0, 1), (-1, -1), 6),
                        
                        # Bordes
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                        ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor('#2c3e50')),
                        
                        # Filas alternadas
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f9f9f9')]),
                        
                        # Ajuste de altura automática para celdas con mucho texto
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ]))
                    
                    # Asegurar que la tabla completa se mantenga junta
                    elements.append(KeepTogether(guias_table))
                
                # Página de firmas
                elements.append(PageBreak())
                elements.append(Spacer(1, 1.5*inch))
                
                firma_data = [
                    ['', '', '', ''],
                    ['_'*35, '', '', '_'*35],
                    ['Entregado por', '', '', 'Recibido por'],
                    ['', '', '', ''],
                    ['Nombre:', '', '', 'Nombre:'],
                    ['', '', '', ''],
                    ['Fecha:', '', '', 'Fecha:'],
                    ['', '', '', ''],
                    ['Hora:', '', '', 'Hora:'],
                ]
                
                firma_table = Table(firma_data, colWidths=[2.8*inch, 1.2*inch, 1.2*inch, 2.8*inch])
                firma_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 2), (-1, 2), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 11),
                ]))
                
                elements.append(firma_table)
                elements.append(Spacer(1, 0.8*inch))
                
                nota_style = ParagraphStyle('Nota', parent=styles['Normal'], 
                    fontSize=9, textColor=colors.HexColor('#666666'), 
                    alignment=TA_CENTER, fontName='Helvetica-Oblique')
                elements.append(Paragraph("Documento generado automáticamente.", nota_style))
                
                # Generar PDF
                doc.build(elements)
                
                # Preparar para descarga
                buffer.seek(0)
                pdf_data = buffer.getvalue()
                
                # Mostrar estadísticas
                st.success("✅ PDF generado exitosamente!")
                st.info(f"""
                **📊 Resumen:**
                - Total órdenes: **{total_ordenes}**
                - Páginas: **{num_paginas + 1}** ({num_paginas} datos + 1 firmas)
                - Fecha: **{FECHA_MANIFIESTO}**
                """)
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar PDF",
                    data=pdf_data,
                    file_name=nombre_pdf,
                    mime="application/pdf",
                    use_container_width=True
                )
                
                # Vista previa del PDF
                with st.expander("👁️ Vista previa del PDF"):
                    pdf_base64 = base64.b64encode(pdf_data).decode()
                    pdf_display = f'<iframe src="data:application/pdf;base64,{pdf_base64}" width="100%" height="500"></iframe>'
                    st.markdown(pdf_display, unsafe_allow_html=True)
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

# Pie de página
st.markdown("---")
st.markdown("🛠️ *Generador de Manifiestos - Automatización Logística*")
