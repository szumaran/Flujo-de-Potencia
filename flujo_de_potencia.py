import streamlit as st
import pandas as pd
import io
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from html2image import Html2Image

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="wide")

# Inicializar html2image de manera segura para entornos Linux/Nube
hti = Html2Image(custom_flags=['--no-sandbox', '--disable-gpu', '--headless'])

# --- 2. MOTOR DE CAPTURA VISUAL (REPLICA COPYPICTURE EN LINUX) ---
def renderizar_bloque_a_imagen(df, keyword):
    """
    Busca el bloque horizontal en el Excel, genera un HTML idéntico a una tabla
    de Excel con tus estilos corporativos, y le toma una captura de pantalla en la nube.
    """
    headers = df.columns.astype(str).tolist()
    columnas_coincidentes = [col for col in headers if keyword in col]
    
    if not columnas_coincidentes:
        return None
        
    # Encontrar los límites horizontales del bloque (columna vacía)
    idx_inicio = headers.index(columnas_coincidentes[0])
    idx_fin = idx_inicio
    while idx_fin < len(df.columns) and not df.iloc[:, idx_fin].isna().all():
        idx_fin += 1
        
    df_bloque = df.iloc[:, idx_inicio:idx_fin].dropna(how='all').reset_index(drop=True)
    if df_bloque.empty:
        return None

    # Reemplazar saltos de línea para el formato web
    df_bloque.columns = [col.replace('\n', '<br>') for col in df_bloque.columns]

    # CSS Estilo Planilla Excel Corporativa (Azul #004C5F)
    html_style = """
    <style>
        table {
            border-collapse: collapse;
            font-family: 'Ubuntu', 'Segoe UI', Arial, sans-serif;
            font-size: 12px;
            width: auto;
            margin: 5px;
        }
        th {
            background-color: #004C5F;
            color: white;
            font-weight: bold;
            border: 1px solid #a0a0a0;
            padding: 6px 12px;
            text-align: center;
            white-space: nowrap;
        }
        td {
            border: 1px solid #d3d3d3;
            padding: 5px 10px;
            color: #004C5F;
            text-align: left;
            white-space: nowrap;
        }
        tr:nth-child(even) {
            background-color: #fcfcfc;
        }
    </style>
    """
    
    html_tabla = df_bloque.to_html(index=False, escape=False)
    html_completo = f"<html><head>{html_style}</head><body style='background-color:white; padding:10px;'>{html_tabla}</body></html>"
    
    nombre_imagen = f"tabla_{keyword.replace(chr(10), '_')}.png"
    
    # Calcular dimensiones dinámicas óptimas para evitar cortes en el recorte de la foto
    ancho_estimado = max(len(df_bloque.columns) * 140, 400)
    alto_estimado = (len(df_bloque) * 28) + 80
    
    try:
        hti.screenshot(html_str=html_completo, save_as=nombre_imagen, size=(ancho_estimado, alto_estimado))
        return nombre_imagen
    except Exception:
        return None

# --- 3. INTERFAZ DE USUARIO ---
st.title("⚡ Plataforma Flujo de Potencia")
st.write("Generador automático de reportes. Extrae rangos horizontales y los pega en Word como capturas de alta definición.")
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
        
        # Formato de texto base
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Ubuntu'
        font.size = Pt(18)
        font.color.rgb = RGBColor(0, 76, 95)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_sheets = len(sheet_names)
        
        # Mapeo de nombres como los tenías en tu macro original
        titleMapping = {
            "Línea\n[MVA]": "Cargabilidad Líneas (MVA)",
            "Línea\n[kA]": "Cargabilidad Líneas (kA)",
            "Transformador": "Cargabilidad Transformadores",
            "Barra": "Regulación de tensión"
        }

        for index, sheet_name in enumerate(sheet_names):
            status_text.write(f"Capturando tablas del Escenario: **{sheet_name}** ({index + 1}/{total_sheets})...")
            
            # Título principal del escenario en Word (Líneas 5-6 de tu VBA)
            p_title = doc.add_paragraph()
            run_title = p_title.add_run(f"Resultados {sheet_name}")
            run_title.font.size = Pt(18)
            run_title.font.bold = True
            run_title.font.name = 'Ubuntu'
            run_title.font.color.rgb = RGBColor(0, 76, 95)
            
            pPr = p_title._p.get_or_add_pPr()
            pPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>'))
            
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Buscar cada una de tus tablas
            for kw, titulo_tabla in titleMapping.items():
                img_file = renderizar_bloque_a_imagen(df_sheet, kw)
                
                if img_file and os.path.exists(img_file):
                    # Título de la tabla en estilo Caption arriba de la imagen
                    p_cap = doc.add_paragraph()
                    run_cap = p_cap.add_run(f"Tabla: Resultados {titulo_tabla} - {sheet_name}")
                    run_cap.font.name = 'Ubuntu'
                    run_cap.font.size = Pt(11)
                    run_cap.font.bold = True
                    run_cap.font.color.rgb = RGBColor(0, 76, 95)
                    
                    # Pegar la imagen fija en el documento Word
                    doc.add_picture(img_file)
                    doc.add_paragraph() # Espacio inferior
                    
                    # Borrar archivo temporal de imagen de la memoria del servidor
                    os.remove(img_file)
            
            if index < total_sheets - 1:
                doc.add_page_break()
                
            progress_bar.progress((index + 1) / total_sheets)
            
        # Preparar descarga
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        status_text.success("✨ ¡El reporte con tus capturas fijas ha sido generado con éxito en la nube!")
        
        st.download_button(
            label="📥 Descargar Informe Word Oficial (.docx)",
            data=docx_buffer,
            file_name="Informe_Flujo_de_Potencia_Oficial.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("---")
st.caption("© 2026 Plataforma Flujo de Potencia - Ingeniería Eléctrica")
