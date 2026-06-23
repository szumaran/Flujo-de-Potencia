import streamlit as st
from docx import Document
import google.generativeai as genai
import io

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="wide")

# --- 2. CONFIGURACIÓN DE LA IA (GOOGLE GEMINI) ---
if "GEMINI_API_KEY" in st.secrets:
    try:
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
        # Usamos tu modelo Gemini de 2026
        model = genai.GenerativeModel('gemini-3-flash-preview')
    except Exception as e:
        st.error(f"Error al configurar la IA: {e}")
else:
    st.error("❌ Falta GEMINI_API_KEY en los Secrets de Streamlit.")

def analizar_escenario_con_ia(nombre_escenario, texto_tablas):
    """Envía los datos de las tablas del escenario a Gemini"""
    prompt = f"""
    Actúa como un Ingeniero Senior de Planificación de Sistemas Eléctricos de Potencia.
    Analiza el comportamiento operativo del escenario '{nombre_escenario}' basado en los datos de sus tablas de resultados:
    
    {texto_tablas}
    
    Genera un informe técnico estructurado para este escenario con dos secciones claras:
    1. ANÁLISIS TÉCNICO COMPLETO: Evalúa la cargabilidad de líneas (MVA y kA), transformadores y perfiles de voltaje en barras (p.u.). Sé específico si detectas anomalías o instalaciones críticas.
    2. CONCLUSIÓN DEL ESCENARIO: Resume de forma ejecutiva la condición operativa del sistema en este caso.
    
    Escribe directamente el análisis de forma fluida, profesional y formal. No agregues introducciones ni saludos.
    """
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Error en el análisis de IA: {str(e)}"

def procesar_documento_online(docx_file):
    doc = Document(docx_file)
    escenario_actual = None
    datos_acumulados = []
    parrafos_a_insertar = []

    # 1. Escanear párrafos buscando títulos de escenarios
    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto.startswith("Resultados Escenario:"):
            if escenario_actual and datos_acumulados:
                analisis = analizar_escenario_con_ia(escenario_actual, "\n".join(datos_acumulados))
                parrafos_a_insertar.append((escenario_actual, analisis))
            
            escenario_actual = texto.replace("Resultados Escenario:", "").strip()
            datos_acumulados = []
        elif escenario_actual:
            datos_acumulados.append(texto)

    # 2. Leer las tablas nativas de Word para la IA
    for table in doc.tables:
        texto_tabla = []
        for row in table.rows:
            fila = [cell.text.strip() for cell in row.cells]
            texto_tabla.append(", ".join(fila))
        if escenario_actual:
            datos_acumulados.append("\n".join(texto_tabla))

    # Procesar el último caso
    if escenario_actual and datos_acumulados:
        analisis = analizar_escenario_con_ia(escenario_actual, "\n".join(datos_acumulados))
        parrafos_a_insertar.append((escenario_actual, analisis))

    # 3. Inyectar los textos debajo de cada escenario
    for esc_nombre, analisis_texto in parrafos_a_insertar:
        for i in range(len(doc.paragraphs) - 1, -1, -1):
            if esc_nombre in doc.paragraphs[i].text:
                p_analisis = doc.paragraphs[i].insert_paragraph_before()
                p_analisis.text = f"\nAnálisis Técnico de IA - Escenario {esc_nombre}:\n{analisis_texto}\n"
                break

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

# --- 3. INTERFAZ WEB ---
st.title("⚡ Analizador Remoto de Flujos de Potencia")
st.write("Potenciado por Google Gemini (Librería nativa)")
st.markdown("---")

uploaded_file = st.file_uploader("Carga tu archivo Word (.docx)", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 Iniciar Análisis Completo"):
        with st.spinner("Gemini está analizando los escenarios eléctricos..."):
            try:
                word_comentado = procesar_documento_online(uploaded_file)
                st.success("¡Informe procesado con éxito!")
                
                st.download_button(
                    label="📥 Descargar Word con Conclusiones de IA",
                    data=word_comentado,
                    file_name="Reporte_Final_Conclusiones_IA.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error en el servidor: {str(e)}")
