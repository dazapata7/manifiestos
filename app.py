import streamlit as st
import pandas as pd
import base64
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.platypus.flowables import KeepInFrame

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
                
                # Tamaño de página landscape optimizado
                doc = SimpleDocTemplate(
                    buffer, 
                    pagesize=landscape(letter),
                    rightMargin=0.5*cm,
                    leftMargin=0.5*cm,
                    topMargin=1.5*cm,
                    bottomMargin=1.0*cm
                )
                
                elements = []
                styles = getSampleStyleSheet()
                
                # Estilos personalizados
                title_style = ParagraphStyle(
                    'CustomTitle', 
                    parent=styles['Heading1'],
                    fontSize=14,
                    textColor=colors.HexColor('#1a1a1a'),
                    spaceAfter=6,
                    alignment=TA_CENTER,
                    fontName='Helvetica-Bold'
                )
                
                subtitle_style = ParagraphStyle(
                    'CustomSubtitle', 
                    parent=styles['Normal'],
                    fontSize=9,
                    textColor=colors.HexColor('#666666'),
                    spaceAfter=10,
                    alignment=TA_CENTER,
                    fontName='Helvetica'
                )
                
                # Estilo para texto en celdas (más compacto)
                cell_style = ParagraphStyle(
                    'CellStyle',
                    parent=styles['Normal'],
                    fontSize=7.5,
                    alignment=TA_LEFT,
                    fontName='Helvetica',
                    leading=9  # Espaciado entre líneas reducido
                )
                
                # Estilo para números
                number_style = ParagraphStyle(
                    'NumberStyle',
                    parent=styles['Normal'],
                    fontSize=8,
                    alignment=TA_CENTER,
                    fontName='Helvetica',
                    leading=9
                )
                
                # ANCHOS DE COLUMNAS OPTIMIZADOS (en cm para precisión)
                col_widths = [
                    0.6*cm,    # # (muy estrecho)
                    1.8*cm,    # Guía
                    3.5*cm,    # Cliente
                    2.0*cm,    # Ciudad
                    2.0*cm,    # Estado
                    3.5*cm,    # Dirección
                    4.0*cm     # Producto
                ]
                
                # Calcular páginas necesarias (18 órdenes por página)
                total_ordenes = len(df)
                ordenes_por_pagina = 18
                num_paginas = (total_ordenes + ordenes_por_pagina - 1) // ordenes_por_pagina
                
                # Generar cada página
                for page_num in range(num_paginas):
                    start_idx = page_num * ordenes_por_pagina
                    end_idx = min((page_num + 1) * ordenes_por_pagina, total_ordenes)
                    
                    if page_num > 0:
                        elements.append(PageBreak())
                    
                    # ENCABEZADO DE PÁGINA
                    elements.append(Paragraph("MANIFIESTO DE ENTREGA", title_style))
                    
                    if num_paginas > 1:
                        elements.append(Paragraph(
                            f"Fecha: {FECHA_MANIFIESTO} | Total: {total_ordenes} órdenes | Página {page_num + 1} de {num_paginas}",
                            subtitle_style
                        ))
                    else:
                        elements.append(Paragraph(
                            f"Fecha: {FECHA_MANIFIESTO} | Total: {total_ordenes} órdenes",
                            subtitle_style
                        ))
                    
                    # Espacio entre encabezado y tabla
                    elements.append(Spacer(1, 0.2*cm))
                    
                    # Obtener datos para esta página
                    chunk = df.iloc[start_idx:end_idx].copy()
                    
                    # Preparar datos de la tabla
                    table_data = []
                    
                    # ENCABEZADOS DE TABLA
                    header_row = [
                        Paragraph('#', cell_style),
                        Paragraph('Guía', cell_style),
                        Paragraph('Cliente', cell_style),
                        Paragraph('Ciudad', cell_style),
                        Paragraph('Estado', cell_style),
                        Paragraph('Dirección', cell_style),
                        Paragraph('Producto', cell_style)
                    ]
                    table_data.append(header_row)
                    
                    # DATOS DE LAS ÓRDENES
                    for idx, row in chunk.iterrows():
                        # Preparar cada campo con límites de caracteres
                        guia = str(row['Guía de Envío'])[:10] if pd.notna(row['Guía de Envío']) else 'N/A'
                        cliente = str(row['Cliente'])[:25] if pd.notna(row['Cliente']) else 'N/A'
                        ciudad = str(row['Ciudad'])[:15] if pd.notna(row['Ciudad']) else 'N/A'
                        estado = str(row['Estado'])[:12] if pd.notna(row['Estado']) else 'N/A'
                        
                        # Dirección
                        direccion_parts = []
                        if pd.notna(row['Calle']):
                            direccion_parts.append(str(row['Calle']))
                        if pd.notna(row['Número']):
                            direccion_parts.append(' ' + str(row['Número']))
                        direccion = ' '.join(direccion_parts)[:35] if direccion_parts else 'N/A'
                        
                        # Producto (con wrap automático)
                        if pd.notna(row['Productos']):
                            producto_text = str(row['Productos'])
                            # Limitar longitud pero permitir wrap
                            if len(producto_text) > 50:
                                producto_text = producto_text[:47] + '...'
                        else:
                            producto_text = 'N/A'
                        
                        # Número de orden CORREGIDO (continuo entre páginas)
                        orden_num = start_idx + (idx - chunk.index[0]) + 1
                        
                        # Crear fila con Paragraphs para mejor control
                        row_data = [
                            Paragraph(str(orden_num), number_style),
                            Paragraph(guia, number_style),
                            Paragraph(cliente, cell_style),
                            Paragraph(ciudad, cell_style),
                            Paragraph(estado, cell_style),
                            Paragraph(direccion, cell_style),
                            Paragraph(producto_text, cell_style)
                        ]
                        
                        table_data.append(row_data)
                    
                    # Crear tabla con alturas de fila reducidas
                    tabla = Table(
                        table_data, 
                        colWidths=col_widths,
                        repeatRows=1,  # Repetir encabezados en cada página
                        rowHeights=0.5*cm  # Altura fija y reducida para todas las filas
                    )
                    
                    # APLICAR ESTILOS A LA TABLA
                    estilo_tabla = TableStyle([
                        # ENCABEZADO
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 8),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                        ('TOPPADDING', (0, 0), (-1, 0), 4),
                        
                        # LÍNEA INFERIOR DEL ENCABEZADO
                        ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.black),
                        
                        # DATOS
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                        ('ALIGN', (0, 1), (1, -1), 'CENTER'),  # Columnas # y Guía centradas
                        ('ALIGN', (2, 1), (-1, -1), 'LEFT'),   # Demás columnas alineadas a la izquierda
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 7.5),
                        ('BOTTOMPADDING', (0, 1), (-1, -1), 2),
                        ('TOPPADDING', (0, 1), (-1, -1), 2),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        
                        # BORDES DELGADOS
                        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),  # Líneas más delgadas
                        
                        # FILAS ALTERNADAS (solo para datos, no encabezado)
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f5f5f5')]),
                    ])
                    
                    tabla.setStyle(estilo_tabla)
                    elements.append(tabla)
                
                # PÁGINA DE FIRMAS (solo una vez al final)
                elements.append(PageBreak())
                elements.append(Spacer(1, 3*cm))
                
                # Tabla de firmas más compacta
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
                
                firma_table = Table(
                    firma_data, 
                    colWidths=[4*cm, 1.5*cm, 1.5*cm, 4*cm],
                    rowHeights=0.6*cm
                )
                
                firma_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 2), (-1, 2), 'Helvetica-Bold'),
                    ('FONTNAME', (0, 4), (0, 4), 'Helvetica-Bold'),
                    ('FONTNAME', (3, 4), (3, 4), 'Helvetica-Bold'),
                    ('FONTNAME', (0, 6), (0, 6), 'Helvetica-Bold'),
                    ('FONTNAME', (3, 6), (3, 6), 'Helvetica-Bold'),
                    ('FONTNAME', (0, 8), (0, 8), 'Helvetica-Bold'),
                    ('FONTNAME', (3, 8), (3, 8), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                ]))
                
                elements.append(firma_table)
                elements.append(Spacer(1, 1.5*cm))
                
                # Nota al pie
                nota_style = ParagraphStyle(
                    'Nota', 
                    parent=styles['Normal'],
                    fontSize=8,
                    textColor=colors.HexColor('#666666'),
                    alignment=TA_CENTER,
                    fontName='Helvetica-Oblique'
                )
                elements.append(Paragraph("Documento generado automáticamente.", nota_style))
                
                # CONSTRUIR EL PDF
                doc.build(elements)
                
                # Preparar archivo para descarga
                buffer.seek(0)
                pdf_data = buffer.getvalue()
                
                # Mostrar estadísticas
                st.success("✅ PDF generado exitosamente!")
                st.info(f"""
                **📊 Resumen:**
                - Total órdenes: **{total_ordenes}**
                - Páginas: **{num_paginas + 1}** ({num_paginas} de datos + 1 de firmas)
                - Fecha: **{FECHA_MANIFIESTO}**
                - Diseño optimizado para 18 órdenes por página
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
        st.exception(e)

# Pie de página
st.markdown("---")
st.markdown("🛠️ *Generador de Manifiestos - Automatización Logística*")
