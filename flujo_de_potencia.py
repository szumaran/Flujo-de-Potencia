import streamlit as st
from docx import Document
from openai import OpenAI
import io

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="centered")

# Pega tu API Key de OpenAI aquí
client = OpenAI(api_key="TU_API_KEY_AQUI")

def analizar_escenario_con_ia(nombre_escenario, texto_tablas):
    """Envía los datos de las tablas del escenario a la IA para el análisis técnico"""
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
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error en el análisis de IA: {str(e)}"

def procesar_documento_online(docx_file):
    # Cargar el archivo Word en la memoria del servidor Linux
    doc = Document(docx_file)
    
    escenario_actual = None
    datos_acumulados = []
    parrafos_a_insertar = []

    # 1. Escanear párrafos en busca de los títulos de tus escenarios
    for p in doc.paragraphs:
        texto = p.text.strip()
        if texto.startswith("Resultados Escenario:"):
            # Si ya teníamos datos guardados de un escenario previo, lo mandamos a la IA
            if escenario_actual and datos_acumulados:
                analisis = analizar_escenario_con_ia(escenario_actual, "\n".join(datos_acumulados))
                parrafos_a_insertar.append((escenario_actual, analisis))
            
            escenario_actual = texto.replace("Resultados Escenario:", "").strip()
            datos_acumulados = []
        elif escenario_actual:
            datos_acumulados.append(texto)

    # 2. Leer los datos de las tablas nativas de Word para dárselos a la IA
    for table in doc.tables:
        texto_tabla = []
        for row in table.rows:
            fila = [cell.text.strip() for cell in row.cells]
            texto_tabla.append(", ".join(fila))
        if escenario_actual:
            datos_acumulados.append("\n".join(texto_tabla))

    # Guardar el procesamiento del último escenario
    if escenario_actual and datos_acumulados:
        analisis = analizar_escenario_con_ia(escenario_actual, "\n".join(datos_acumulados))
        parrafos_a_insertar.append((escenario_actual, analisis))

    # 3. Inyectar dinámicamente las respuestas de la IA debajo de cada bloque correspondiente
    for esc_nombre, analisis_texto in parrafos_a_insertar:
        for i in range(len(doc.paragraphs) - 1, -1, -1):
            if esc_nombre in doc.paragraphs[i].text:
                p_analisis = doc.paragraphs[i].insert_paragraph_before()
                p_analisis.text = f"\nAnálisis Técnico de IA - Escenario {esc_nombre}:\n{analisis_texto}\n"
                break

    # Guardar los cambios aplicados en el Word en memoria binaria
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

# --- 3. INTERFAZ EN LA NUBE ---
st.title("⚡ Analizador Remoto de Flujos de Potencia")
st.write("Sube el documento Word (.docx) que generó tu macro para inyectarle los comentarios y conclusiones automáticas de la IA.")
st.markdown("---")

uploaded_file = st.file_uploader("Carga tu archivo Word (.docx)", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 Iniciar Análisis Completo"):
        with st.spinner("La IA de la nube está analizando los escenarios eléctricos..."):
            try:
                word_comentado = procesar_documento_online(uploaded_file)
                st.success("¡Análisis inyectado con éxito!")
                
                st.download_button(
                    label="📥 Descargar Word con Conclusiones de IA",
                    data=word_comentado,
                    file_name="Reporte_Final_Conclusiones_IA.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error en el servidor: {str(e)}")
