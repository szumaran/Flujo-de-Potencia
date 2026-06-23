import streamlit as st
import pandas as pd
import io
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="wide")

# --- 2. LÓGICA DE EXTRACCIÓN HORIZONTAL Y DIBUJO DE TABLAS ---
def extraer_y_dibujar_bloques_reales(df, doc, sheet_name):
    """
    Busca de forma exacta las sub-tablas distribuidas horizontalmente en tu Excel
    y las plasma en Word con el formato visual corporativo limpio.
    """
    azul_corporativo = RGBColor(0, 76, 95)
    
    # Mapeo idéntico de palabras clave y sus títulos de tu macro de VBA
    titleMapping = {
        "Línea\n[MVA]": "Cargabilidad Líneas (MVA)",
        "Línea\n[kA]": "Cargabilidad Líneas (kA)",
        "Transformador": "Cargabilidad Transformadores",
        "Barra": "Regulación de tensión"
    }
    
    # Título principal del Escenario en el Word (Líneas 5-6 de tu VBA original)
    p_title = doc.add_paragraph()
    run_title = p_title.add_run(f"Resultados {sheet_name}")
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.name = 'Ubuntu'
    run_title.font.color.rgb = azul_corporativo
    
    pPr = p_title._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>'))
    
    # Analizar el archivo buscando cada bloque por su columna de inicio
    for kw, titulo_tabla in titleMapping.items():
        # Encontrar los índices de las columnas que contienen la palabra clave
        columnas_bloque = [i for i, col in enumerate(df.columns) if kw in str(col)]
        
        if columnas_bloque:
            idx_inicio = columnas_bloque[0]
            
            # Un bloque dura horizontalmente hasta encontrar una columna vacía (NaN)
            idx_fin = idx_inicio
            while idx_fin < len(df.columns) and not df.iloc[:, idx_fin].isna().all():
                idx_fin += 1
                
            # Extraer el trozo exacto de la tabla horizontal sin mezclar columnas ajenas
            df_sub_tabla = df.iloc[:, idx_inicio:idx_fin].dropna(how='all').reset_index(drop=True)
            
            if not df_sub_tabla.empty:
                # Título estilo CAPTION arriba de la tabla (Exactamente como tu macro)
                p_nom = doc.add_paragraph()
                run_nom = p_nom.add_run(f"Tabla: Resultados {titulo_tabla} - {sheet_name}")
                run_nom.font.name = 'Ubuntu'
                run_nom.font.size = Pt(11)
                run_nom.font.bold = True
                run_nom.font.color.rgb = azul_corporativo
                
                # Crear la tabla física en Word
                tabla_word = doc.add_table(rows=len(df_sub_tabla) + 1, cols=len(df_sub_tabla.columns))
                tabla_word.style = 'Table Grid'
                
                # 1. Escribir los encabezados reales de las columnas
                for j, col_name in enumerate(df_sub_tabla.columns):
                    cell = tabla_word.cell(0, j)
                    cell.text = str(col_name).replace(chr(10), ' ')
                    if cell.paragraphs[0].runs:
                        run = cell.paragraphs[0].runs[0]
                        run.font.name = 'Ubuntu'
                        run.font.size = Pt(10)
                        run.font.bold = True
                        run.font.color.rgb = azul_corporativo
                
                # 2. Escribir las celdas de datos con formato numérico limpio
                for i, row in enumerate(df_sub_tabla.itertuples(index=False)):
                    for j, val in enumerate(row):
                        cell = tabla_word.cell(i + 1, j)
                        
                        # Controlar el redondeo de decimales para evitar deformaciones
                        if isinstance(val, float):
                            if "p.u." in str(df_sub_tabla.columns[j]) or "Tensión" in str(df_sub_tabla.columns[j]):
                                cell.text = f"{val:.1f}"  # 1 decimal para tensiones
                            else:
                                cell.text = f"{val:.2f}"  # 2 decimales para cargabilidades
                        else:
                            cell.text = str(val) if pd.notna(val) else ""
                            
                        if cell.paragraphs[0].runs:
                            run_cell = cell.paragraphs[0].runs[0]
                            run_cell.font.name = 'Ubuntu'
                            run_cell.font.size = Pt(10)
                            run_cell.font.color.rgb = azul_corporativo
                            
                # Espacio de separación después de cada bloque de tabla
                doc.add_paragraph()

# --- 3. INTERFAZ DE USUARIO ---
st.title("⚡ Plataforma Flujo de Potencia")
st.write("Extractor automático de simulación. Organiza los bloques horizontales directamente en tu Word oficial.")
st.markdown("---")

uploaded_file = st.file_uploader("Selecciona el archivo Excel de Resultados (EFP_RES_...)", type=["xlsx"])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        st.success(f"📂 Archivo cargado correctamente. Se detectaron {len(sheet_names)} escenarios de estudio.")
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        st.stop()

    if st.button("🚀 Generar Informe Word Oficial"):
        
        doc = Document()
        
        # Formato de texto base por defecto
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Ubuntu'
        font.size = Pt(18)
        font.color.rgb = RGBColor(0, 76, 95)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_sheets = len(sheet_names)

        # Iterar por cada escenario (hoja del Excel)
        for index, sheet_name in enumerate(sheet_names):
            status_text.write(f"Escribiendo tablas del Escenario: **{sheet_name}** ({index + 1}/{total_sheets})...")
            
            # Leer la hoja actual respetando el formato original
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Ejecutar la separación horizontal exacta de tus bloques
            extraer_y_dibujar_bloques_reales(df_sheet, doc, sheet_name)
            
            if index < total_sheets - 1:
                doc.add_page_break()
                
            progress_bar.progress((index + 1) / total_sheets)
            
        # Empaquetar para la descarga remota en memoria RAM
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        status_text.success("✨ ¡El informe oficial con todas tus tablas se ha compilado con éxito!")
        
        st.download_button(
            label="📥 Descargar Informe Word Oficial (.docx)",
            data=docx_buffer,
            file_name="Informe_Flujo_de_Potencia_Oficial.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("---")
st.caption("© 2026 Plataforma Flujo de Potencia - Ingeniería Eléctrica")
