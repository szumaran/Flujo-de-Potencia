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
    # Usamos tu modelo Gemini de 2026
    model = genai.GenerativeModel('gemini-3-flash-preview')
except Exception as e:
    st.error(f"Error al configurar la IA de Google: {e}")

def analizar_escenario_con_ia(nombre_escenario, texto_tablas):
    prompt = f"""Actúa como un Ingeniero Senior de Planificación de Sistemas Eléctricos de Potencia.
    Analiza el comportamiento operativo del escenario '{nombre_escenario}' basado estrictamente en los siguientes datos de tablas de resultados:
    
    {texto_tablas}
    
    Instrucciones obligatorias de redacción:
    1. Genera un único párrafo continuo de análisis técnico (prosa fluida).
    2. Está estrictamente PROHIBIDO usar listas, viñetas, guiones, puntos apartes o clasificaciones.
    3. Está estrictamente PROHIBIDO usar formato Markdown como negritas (no uses caracteres '**'), títulos o subtítulos (no uses '#').
    4. El texto debe ser plano, directo, formal y puramente técnico.
    5. No redactes ninguna conclusión para este escenario individual (la conclusión general se hará al final del documento).
    6. Destaca únicamente los casos más representativos del escenario: menciona de forma corrida las líneas o transformadores con mayor cargabilidad (MVA, kA) y los perfiles de voltaje en barras (p.u.) más críticos.
    
    Escribe directamente el párrafo de análisis técnico, sin introducciones, saludos, comentarios ni títulos de secciones.
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

    for table in doc.tables:
        texto_tabla = []
        for row in table.rows:
            fila = [cell.text.strip() for cell in row.cells]
            texto_tabla.append(", ".join(fila))
        if escenario_actual:
            datos_acumulados.append("\n".join(texto_tabla))

    if escenario_actual and datos_acumulados:
        analisis = analizar_escenario_con_ia(escenario_actual, "\n".join(datos_acumulados))
        parrafos_a_insertar.append((escenario_actual, analisis))

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

# --- 2. INTERFAZ WEB ---
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
