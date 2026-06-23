import streamlit as st
from docx import Document
import google.generativeai as genai
import io

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="centered")

# Inyección directa de tu clave de Gemini
p1 = "AQ.Ab8RN6LODoM0i-"
p2 = "7K7R7AxOmrhDmJFf_"
p3 = "1ZcyPvbCUFgxaSX5kng"

try:
    genai.configure(api_key=p1 + p2 + p3)
    model = genai.GenerativeModel('gemini-3-flash-preview')
except Exception as e:
    st.error(f"Error al configurar la IA de Google: {e}")

def analizar_escenario_con_ia(nombre_escenario, texto_tablas):
    """Genera UN SOLO párrafo fluido destacando lo más representativo"""
    prompt = f"""
    Actúa como un Ingeniero Senior de Planificación de Sistemas Eléctricos de Potencia.
    Analiza el comportamiento operativo del escenario '{nombre_escenario}' basado en estos datos de tablas:
    
    {texto_tablas}
    
    Genera UN SOLO PÁRRAFO de análisis técnico. Debe ser fluido, continuo, formal y directo, sin usar viñetas, subtítulos ni listas. 
    Destaca únicamente los casos más representativos del escenario: menciona específicamente las líneas o transformadores con mayor cargabilidad (MVA, kA) y los perfiles de voltaje en barras (p.u.) más críticos o relevantes.
    Escribe directamente el párrafo de análisis técnico, sin introducciones ni saludos.
    """
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        return f"Error en el análisis de IA: {str(e)}"

def generar_conclusion_general_ia(texto_completo_documento):
    """Genera una gran conclusión general al final para los 4 escenarios"""
    prompt = f"""
    Actúa como un Ingeniero Senior de Planificación de Sistemas Eléctricos de Potencia.
    Basándote en todos los escenarios analizados en el siguiente documento:
    
    {texto_completo_documento}
    
    Redacta una CONCLUSIÓN GENERAL técnica y ejecutiva que resuma el comportamiento global del sistema a través de los 4 escenarios evaluados. 
    Identifica patrones, instalaciones críticas recurrentes y la robustez general de la red de transmisión frente a los distintos bloques de demanda.
    No uses subtítulos ni viñetas, redacta en párrafos formales de ingeniería.
    """
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        return f"Error al generar la conclusión general: {str(e)}"

def procesar_documento_online(docx_file):
    doc = Document(docx_file)
    escenario_actual = None
    datos_acumulados = []
    parrafos_a_insertar = []
    texto_para_conclusion = ""

    # 1. Agrupar la información por escenarios
    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto.startswith("Resultados Escenario:"):
            if escenario_actual and datos_acumulados:
                analisis = analizar_escenario_con_ia(escenario_actual, "\n".join(datos_acumulados))
                parrafos_a_insertar.append((escenario_actual, analisis))
                texto_para_conclusion += f"\nEscenario {escenario_actual}:\n{analisis}\n"
            
            escenario_actual = texto.replace("Resultados Escenario:", "").strip()
            datos_acumulados = []
        elif escenario_actual:
            datos_acumulados.append(texto)

    for table in doc.tables:
        texto_tabla = []
        for row in table.rows:
            fila = [cell.text.strip() for cell in row.cells]
            texto_tabla.append(", ".join(fila))
        if escenario_actual:
            datos_acumulados.append("\n".join(texto_tabla))

    # Procesar el último escenario detectado
    if escenario_actual and datos_acumulados:
        analisis = analizar_escenario_con_ia(escenario_actual, "\n".join(datos_acumulados))
        parrafos_a_insertar.append((escenario_actual, analisis))
        texto_para_conclusion += f"\nEscenario {escenario_actual}:\n{analisis}\n"

    # 2. Inyectar análisis individuales con formato estricto: Ubuntu 11, Justificado
    for esc_nombre, analisis_texto in parrafos_a_insertar:
        for i in range(len(doc.paragraphs) - 1, -1, -1):
            if esc_nombre in doc.paragraphs[i].text:
                p_analisis = doc.paragraphs[i].insert_paragraph_before()
                p_analisis.text = f"\nAnálisis Técnico de IA - Escenario {esc_nombre}:\n{analisis_texto}\n"
                
                # Forzar formato a nivel de estilo en cada párrafo nuevo
                p_analisis.alignment = 3  # 3 es el código para JUSTIFICADO en python-docx
                for run in p_analisis.runs:
                    run.font.name = 'Ubuntu'
                    run.font.size = io.pt(11)
                break

    # 3. Generar e Inyectar la Conclusión General al final del Word
    conclusion_general_texto = generar_conclusion_general_ia(texto_para_conclusion)
    
    doc.add_heading("CONCLUSIÓN GENERAL DEL ESTUDIO", level=1)
    p_conclusion = doc.add_paragraph(conclusion_general_texto)
    p_conclusion.alignment = 3  # Justificado
    for run in p_conclusion.runs:
        run.font.name = 'Ubuntu'
        run.font.size = io.pt(11)

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

# --- 3. INTERFAZ WEB ---
st.title("⚡ Analizador Remoto de Flujos de Potencia")
st.write("Potenciado por Google Gemini (Formato Corporativo Ubuntu 11)")
st.markdown("---")

uploaded_file = st.file_uploader("Carga tu archivo Word (.docx)", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 Iniciar Análisis Completo"):
        with st.spinner("Gemini está analizando y formateando el informe técnico..."):
            try:
                word_comentado = procesar_documento_online(uploaded_file)
                st.success("¡Informe formateado con éxito!")
                
                st.download_button(
                    label="📥 Descargar Word con Conclusiones de IA",
                    data=word_comentado,
                    file_name="Reporte_Final_Conclusiones_IA.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error en el procesador: {str(e)}")
