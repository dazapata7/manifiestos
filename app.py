import streamlit as st
import pandas as pd
import base64
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
                
                # Configuración del documento
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
                
                # Estilo para título
                title_style = ParagraphStyle(
                    'TitleStyle',
                    parent=styles['Heading1'],
                    fontSize=14,
                    textColor=colors.HexColor('#000000'),
                    alignment=TA_CENTER,
                    spaceAfter=6,
                    fontName='Helvetica-Bold'
                )
                
                # Estilo para subtítulo
                subtitle_style = ParagraphStyle(
                    'SubtitleStyle',
                    parent=styles['Normal'],
                    fontSize=9,
                    textColor=colors.HexColor('#666666'),
                    alignment=TA_CENTER,
                    spaceAfter=12,
                    fontName='Helvetica'
                )
                
                # ANCHOS OPTIMIZADOS (en pulgadas)
                col_widths = [
                    0.4 * inch,   # # (muy estrecho)
                    0.8 * inch,   # Guía (reducido)
                    1.8 * inch,   # Cliente (amplio)
                    1.0 * inch,   # Ciudad
                    1.0 * inch,   # Estado
                    1.9 * inch,   # Dirección
                    2.0 * inch    # Producto (con espacio para texto largo)
                ]
                
                total_ordenes = len(df)
                ordenes_por_pagina = 18
                num_paginas = (total_ordenes + ordenes_por_pagina - 1) // ordenes_por_pagina
                
                # GENERAR PÁGINAS DE DATOS
                for pagina in range(num_paginas):
                    inicio = pagina * ordenes_por_pagina
                    fin = min((pagina + 1) * ordenes_por_pagina, total_ordenes)
                    
                    if pagina > 0:
                        elements.append(PageBreak())
                    
                    # ENCABEZADO DE PÁGINA
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
                    
                    # PREPARAR DATOS DE LA TABLA
                    chunk = df.iloc[inicio:fin]
                    table_data = []
                    
                    # ENCABEZADOS
                    headers = ['#', 'Guía', 'Cliente', 'Ciudad', 'Estado', 'Dirección', 'Producto']
                    table_data.append(headers)
                    
                    # DATOS
                    for idx, row in chunk.iterrows():
                        # Preparar cada valor con límites
                        numero_orden = inicio + (idx - chunk.index[0]) + 1
                        
                        guia = str(row['Guía de Envío'])[:8] if pd.notna(row['Guía de Envío']) else 'N/A'
                        
                        cliente = str(row['Cliente'])[:25] if pd.notna(row['Cliente']) else 'N/A'
                        if len(str(row['Cliente'])) > 25:
                            cliente = cliente[:22] + '...'
                        
                        ciudad = str(row['Ciudad'])[:12] if pd.notna(row['Ciudad']) else 'N/A'
                        estado = str(row['Estado'])[:10] if pd.notna(row['Estado']) else 'N/A'
                        
                        # Dirección combinada
                        direccion_parts = []
                        if pd.notna(row['Calle']):
                            direccion_parts.append(str(row['Calle']))
                        if pd.notna(row['Número']):
                            direccion_parts.append(str(row['Número']))
                        direccion = ' '.join(direccion_parts)
                        if len(direccion) > 30:
                            direccion = direccion[:27] + '...'
                        
                        # Producto - manejo especial para texto largo
                        if pd.notna(row['Productos']):
                            producto = str(row['Productos'])
                            # Si es muy largo, dividirlo en líneas
                            if len(producto) > 35:
                                # Buscar espacio para dividir
                                if ' ' in producto[30:40]:
                                    espacio = producto.find(' ', 30, 40)
                                    if espacio != -1:
                                        producto = producto[:espacio] + '\n' + producto[espacio+1:45] + '...'
                                else:
                                    producto = producto[:35] + '...'
                        else:
                            producto = 'N/A'
                        
                        # Agregar fila
                        table_data.append([
                            str(numero_orden),
                            guia,
                            cliente,
                            ciudad,
                            estado,
                            direccion,
                            producto
                        ])
                    
                    # CREAR TABLA
                    tabla = Table(
                        table_data,
                        colWidths=col_widths,
                        repeatRows=1  # Repetir encabezados
                    )
                    
                    # ESTILOS DE LA TABLA
                    estilo = TableStyle([
                        # ENCABEZADOS
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 9),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                        ('TOPPADDING', (0, 0), (-1, 0), 6),
                        
                        # DATOS
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                        ('ALIGN', (0, 1), (1, -1), 'CENTER'),  # # y Guía centrados
                        ('ALIGN', (2, 1), (-1, -1), 'LEFT'),   # Resto alineado izquierda
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),
                        ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
                        ('TOPPADDING', (0, 1), (-1, -1), 4),
                        
                        # ALTURA DE FILAS FLEXIBLE para producto multilínea
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        
                        # BORDES DELGADOS
                        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
                        
                        # FILAS ALTERNADAS
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
                    ])
                    
                    tabla.setStyle(estilo)
                    elements.append(tabla)
                
                # PÁGINA DE FIRMAS
                elements.append(PageBreak())
                elements.append(Spacer(1, 1.5 * inch))
                
                # Tabla de firmas simple
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
                    colWidths=[3.0 * inch, 0.5 * inch, 0.5 * inch, 3.0 * inch]
                )
                
                firma_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 2), (-1, 2), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 11),
                ]))
                
                elements.append(firma_table)
                elements.append(Spacer(1, 0.8 * inch))
                
                # Nota al pie
                nota_style = ParagraphStyle(
                    'NotaStyle',
                    parent=styles['Normal'],
                    fontSize=9,
                    textColor=colors.HexColor('#666666'),
                    alignment=TA_CENTER,
                    fontName='Helvetica-Oblique'
                )
                elements.append(Paragraph("Documento generado automáticamente.", nota_style))
                
                # GENERAR PDF
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
                - Formato optimizado para 18 órdenes por página
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
        import traceback
        st.code(traceback.format_exc())

# Pie de página
st.markdown("---")
st.markdown("🛠️ *Generador de Manifiestos - Automatización Logística*")
