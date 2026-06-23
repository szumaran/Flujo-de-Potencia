import streamlit as st
import pandas as pd
import io
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import google.generativeai as genai

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="wide")

# --- 2. CONFIGURACIÓN DE LA IA ---
if "GEMINI_API_KEY" in st.secrets:
    try:
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
        model = genai.GenerativeModel('gemini-3-flash-preview')
    except Exception as e:
        st.error(f"Error al configurar la IA: {e}")
else:
    st.error("❌ Falta GEMINI_API_KEY en los Secrets de Streamlit.")
    st.stop()

# --- 3. LÓGICA AUXILIAR DE PROCESAMIENTO ---
def extraer_tablas_de_hoja(df):
    """
    Escanea la hoja (escenario) identificando las columnas y los bloques de datos 
    mediante las palabras clave de potencia, corriente, transformadores y barras.
    """
    resumen_bloques = ""
    keywords = ["Línea\n[MVA]", "Línea\n[kA]", "Transformador", "Barra"]
    
    headers = df.columns.astype(str).tolist()
    
    for keyword in keywords:
        columnas_coincidentes = [col for col in headers if keyword in col]
        
        if columnas_coincidentes:
            resumen_bloques += f"\n### TABLA ENCONTRADA: {keyword.replace(chr(10), ' ')}\n"
            df_bloque = df[columnas_coincidentes].dropna(how='all')
            resumen_bloques += df_bloque.to_markdown(index=False) + "\n"
            
    if not resumen_bloques:
        resumen_bloques = df.dropna(how='all').to_markdown(index=False)
        
    return resumen_bloques

# --- 4. INTERFAZ DE USUARIO ---
st.title("⚡ Plataforma Flujo de Potencia")
st.write("Potenciado por Gemini 3 Flash Preview")
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

    if st.button("🚀 Ejecutar Análisis y Generar Informe"):
        
        doc = Document()
        
        # Configurar formato general solicitado: Fuente Ubuntu, Tamaño 18, Color #004C5F
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Ubuntu'
        font.size = Pt(18)
        font.color.rgb = RGBColor(0, 76, 95)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_sheets = len(sheet_names)
        contexto_para_conclusion = ""

        # Procesar cada hoja del Excel (Escenarios independientes)
        for index, sheet_name in enumerate(sheet_names):
            status_text.write(f"Procesando Escenario: **{sheet_name}** ({index + 1}/{total_sheets})...")
            
            # Título de la sección en Word
            p_title = doc.add_paragraph()
            run_title = p_title.add_run(f"Resultados {sheet_name}")
            run_title.font.size = Pt(18)
            run_title.font.bold = True
            run_title.font.name = 'Ubuntu'
            run_title.font.color.rgb = RGBColor(0, 76, 95)
            
            pPr = p_title._p.get_or_add_pPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>')
            pPr.append(shd)
            
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            contexto_tablas = extraer_tablas_de_hoja(df_sheet)
            
            # Prompt de auditoría eléctrica adaptado a la regulación chilena
            prompt_escenario = f"""
            Actúa como un Ingeniero Consultor Senior en Sistemas Eléctricos de Potencia en Chile.
            Analiza rigurosamente los siguientes resultados de flujo de potencia del escenario '{sheet_name}'.
            
            Información técnica del escenario (Markdown):
            {contexto_tablas}
            
            Redacta un análisis técnico estructurado y ejecutivo (máximo 2 párrafos) sobre la condición operativa de este escenario.
            Criterios normativos a verificar en los datos:
            1. Cargabilidad de Líneas (tanto en la tabla de MVA como en la de kA) y Transformadores: Deben ser inferiores o iguales al 100%. Reporta cualquier sobrecarga si los porcentajes superan este límite.
            2. Regulación de Tensión en Barras: En sistemas menores a 200 kV (como las redes de 110 kV presentes), la tensión bajo régimen estacionario (p.u.) debe mantenerse estrictamente entre 0.93 p.u. y 1.07 p.u.
            
            Si todas las variables se encuentran en rangos de operation segura, confírmalo formalmente.
            """
            
            try:
                response = model.generate_content(prompt_escenario)
                analisis_texto = response.text
            except Exception as e:
                analisis_texto = f"Error al generar el análisis automático: {e}"
            
            # Escribir evaluación en el archivo Word
            p_analisis = doc.add_paragraph()
            run_analisis = p_analisis.add_run(analisis_texto)
            run_analisis.font.name = 'Ubuntu'
            run_analisis.font.size = Pt(18)
            run_analisis.font.color.rgb = RGBColor(0, 76, 95)
            p_analisis.paragraph_format.space_after = Pt(18)
            
            pPr_a = p_analisis._p.get_or_add_pPr()
            pPr_a.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>'))
            
            contexto_para_conclusion += f"\n--- Escenario {sheet_name} ---\n{analisis_texto}\n"
            
            progress_bar.progress((index + 1) / total_sheets)
        
        # Conclusiones Generales
        status_text.write("Redactando Conclusión General del Estudio...")
        
        prompt_conclusion = f"""
        Tomando en cuenta los resultados de todas las contingencias y escenarios analizados anteriormente, 
        redacta la sección de 'Conclusiones y Recomendaciones' definitiva para este informe de flujo de potencia.
        
        Historial de evaluaciones por escenario:
        {contexto_para_conclusion}
        
        Entrega una respuesta con un tono altamente corporativo, formal y de nivel auditoría de ingeniería.
        """
        
        try:
            response_conclusion = model.generate_content(prompt_conclusion)
            conclusion_texto = response_conclusion.text
        except Exception as e:
            conclusion_texto = f"No se pudo compilar la conclusión automáticamente: {e}"
            
        # Añadir bloque final de Conclusiones al Word
        doc.add_page_break()
        p_c_title = doc.add_paragraph()
        run_c_title = p_c_title.add_run("Conclusiones y Recomendaciones Generales")
        run_c_title.font.size = Pt(18)
        run_c_title.font.bold = True
        run_c_title.font.name = 'Ubuntu'
        run_c_title.font.color.rgb = RGBColor(0, 76, 95)
        
        pPr_c = p_c_title._p.get_or_add_pPr()
        pPr_c.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>'))
        
        p_c_text = doc.add_paragraph()
        run_c_text = p_c_text.add_run(conclusion_texto)
        run_c_text.font.name = 'Ubuntu'
        run_c_text.font.size = Pt(18)
        run_c_text.font.color.rgb = RGBColor(0, 76, 95)
        
        pPr_ct = p_c_text._p.get_or_add_pPr()
        pPr_ct.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>'))
        
        # Guardar en memoria para descarga remota
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        status_text.success("✨ ¡Informe técnico de Flujo de Potencia generado con éxito!")
        
        st.download_button(
            label="📥 Descargar Informe Técnico (.docx)",
            data=docx_buffer,
            file_name="Informe_Tecnico_Flujo_de_Potencia.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("---")
st.caption("© 2026 Plataforma Flujo de Potencia - Ingeniería Eléctrica")
