import streamlit as st
import pandas as pd
import io
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="wide")

# --- 2. LÓGICA DE EXTRACCIÓN Y CONSTRUCCIÓN DE TABLAS NATIVAS ---
def procesar_y_dibujar_tablas(df, doc):
    """
    Busca los bloques de datos en la hoja de Excel según tus encabezados reales
    y los dibuja uno a uno de forma nativa en el documento Word.
    """
    keywords = ["Línea\n[MVA]", "Línea\n[kA]", "Transformador", "Barra"]
    headers = df.columns.astype(str).tolist()
    azul_corporativo = RGBColor(0, 76, 95)
    
    tablas_procesadas = 0
    
    for keyword in keywords:
        # Identificar las columnas que pertenecen a este bloque
        columnas_coincidentes = [col for col in headers if keyword in col]
        
        if columnas_coincidentes:
            tablas_procesadas += 1
            nombre_tabla = keyword.replace(chr(10), ' ')
            
            # Filtrar el sub-dataframe de la tabla eliminando filas completamente vacías
            df_bloque = df[columnas_coincidentes].dropna(how='all')
            
            # Título de la Tabla en el Word
            p_nom_tabla = doc.add_paragraph()
            run_nt = p_nom_tabla.add_run(f"Tabla: {nombre_tabla}")
            run_nt.font.name = 'Ubuntu'
            run_nt.font.size = Pt(14)
            run_nt.font.bold = True
            run_nt.font.color.rgb = azul_corporativo
            
            # Crear estructura de la tabla nativa (filas = datos + cabecera, columnas)
            tabla_word = doc.add_table(rows=len(df_bloque) + 1, cols=len(df_bloque.columns))
            tabla_word.style = 'Table Grid'
            
            # 1. Escribir los encabezados en la primera fila de la tabla
            for j, col_name in enumerate(df_bloque.columns):
                cell = tabla_word.cell(0, j)
                cell.text = str(col_name).replace(chr(10), ' ')
                # Formato del encabezado de la tabla
                if cell.paragraphs[0].runs:
                    run = cell.paragraphs[0].runs[0]
                    run.font.name = 'Ubuntu'
                    run.font.size = Pt(12)
                    run.font.bold = True
                    run.font.color.rgb = azul_corporativo
                
            # 2. Escribir las filas con los datos técnicos
            for i, row in enumerate(df_bloque.itertuples(index=False)):
                for j, val in enumerate(row):
                    cell = tabla_word.cell(i + 1, j)
                    
                    # Formatear el despliegue numérico
                    if isinstance(val, float):
                        # Si es columna de tensión o p.u., dejar con 1 decimal como pediste
                        if "Tensión" in df_bloque.columns[j] or "p.u." in df_bloque.columns[j]:
                            cell.text = f"{val:.1f}"
                        else:
                            cell.text = f"{val:.2f}"
                    else:
                        cell.text = str(val) if pd.notna(val) else ""
                    
                    # Formato de las celdas de datos
                    if cell.paragraphs[0].runs:
                        run_cell = cell.paragraphs[0].runs[0]
                        run_cell.font.name = 'Ubuntu'
                        run_cell.font.size = Pt(11)
                        run_cell.font.color.rgb = azul_corporativo
            
            # Añadir un espacio de separación después de la tabla
            doc.add_paragraph()
            
    return tablas_procesadas

# --- 3. INTERFAZ DE USUARIO ---
st.title("⚡ Plataforma Flujo de Potencia")
st.write("Estructuración y generación nativa de reportes técnicos de simulación.")
st.markdown("---")

uploaded_file = st.file_uploader("Selecciona el archivo Excel de Resultados (EFP_RES_...)", type=["xlsx"])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        st.success(f"📂 Archivo cargado correctamente. Se detectaron {len(sheet_names)} escenarios (hojas) para procesar.")
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        st.stop()

    if st.button("🚀 Generar Tablas en Documento Word"):
        
        doc = Document()
        
        # Configurar formato base del documento: Ubuntu, Tamaño 18, Color Azul Oscuro #004C5F
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Ubuntu'
        font.size = Pt(18)
        font.color.rgb = RGBColor(0, 76, 95)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_sheets = len(sheet_names)

        # Recorrer cada hoja (Escenario) del archivo de entrada
        for index, sheet_name in enumerate(sheet_names):
            status_text.write(f"Dibujando tablas del Escenario: **{sheet_name}** ({index + 1}/{total_sheets})...")
            
            # Título de la sección del Escenario
            p_title = doc.add_paragraph()
            run_title = p_title.add_run(f"Resultados {sheet_name}")
            run_title.font.size = Pt(18)
            run_title.font.bold = True
            run_title.font.name = 'Ubuntu'
            run_title.font.color.rgb = RGBColor(0, 76, 95)
            
            # Forzar fondo blanco
            pPr = p_title._p.get_or_add_pPr()
            pPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>'))
            
            # Leer los datos de la hoja actual
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Llamar a la función para incrustar las tablas del escenario actual
            procesar_y_dibujar_tablas(df_sheet, doc)
            
            # Separador estético antes del siguiente escenario
            p_sep = doc.add_paragraph()
            p_sep.paragraph_format.space_after = Pt(24)
            
            progress_bar.progress((index + 1) / total_sheets)
        
        # Guardar el archivo estructurado en la memoria RAM para la descarga
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        status_text.success("✨ ¡El documento con las tablas nativas se ha generado correctamente!")
        
        st.download_button(
            label="📥 Descargar Informe con Tablas (.docx)",
            data=docx_buffer,
            file_name="Informe_Estructurado_Flujo_de_Potencia.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("---")
st.caption("© 2026 Plataforma Flujo de Potencia - Ingeniería Eléctrica")
