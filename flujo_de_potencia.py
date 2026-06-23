import streamlit as st
import pandas as pd
import os
import win32com.client as win32

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="wide")

st.title("⚡ Plataforma Flujo de Potencia (Modo Local Windows)")
st.write("Busca los bloques del Excel, les toma una captura y los pega en Word con su respectivo título arriba (Caption).")
st.markdown("---")

# Input para definir la ruta de tu computador donde se guardará el Word final
ruta_destino = st.text_input("📁 Carpeta local donde se guardará el informe Word:", value="C:\\Users\\Public\\Documents")
uploaded_file = st.file_uploader("Selecciona el archivo Excel de Resultados (EFP_RES_...)", type=["xlsx", "xlsm"])

def ejecutar_traspaso_imagenes_vba(excel_path, output_dir):
    """
    Réplica exacta de tu macro de VBA. Busca las palabras clave, copia el rango 
    como imagen fija (.CopyPicture) y lo pega en Word agregando el Caption de título arriba.
    """
    # Constantes nativas de la API de Microsoft Office para Python
    xlValues = -4163
    xlPart = 2
    xlByRows = 1
    xlNext = 1
    xlScreen = 1
    xlPicture = -4147
    wdCaptionPositionAbove = 0
    
    # 1. Levantar Word en segundo plano
    wordApp = win32.DispatchEx("Word.Application")
    wordApp.Visible = False
    wordDoc = wordApp.Documents.Add()
    
    # Asegurar que exista el rótulo "Tabla" para el Caption
    lbl_existe = False
    for lbl in wordApp.CaptionLabels:
        if lbl.Name == "Tabla":
            lbl_existe = True
            break
    if not lbl_existe:
        wordApp.CaptionLabels.Add(Name="Tabla")
        
    # 2. Levantar Excel en segundo plano
    excelApp = win32.DispatchEx("Excel.Application")
    excelApp.Visible = False
    excelWorkbook = excelApp.Workbooks.Open(excel_path)
    
    # Diccionario con los títulos que tu macro le asigna a cada bloque
    titleMapping = {
        "Línea\n[MVA]": "Cargabilidad Líneas (MVA)",
        "Línea\n[kA]": "Cargabilidad Líneas (kA)",
        "Transformador": "Cargabilidad Transformadores",
        "Barra": "Regulación de tensión"
    }
    
    # Recorrer cada hoja (Escenario) tal cual lo hace tu bucle de VBA
    for excelSheet in excelWorkbook.Sheets:
        
        # Insertar título del caso (Líneas 5-6 de tu VBA original)
        p_title = wordDoc.Content.Paragraphs.Add().Range
        p_title.Text = f"Resultados {excelSheet.Name}"
        p_title.Font.Name = "Ubuntu"
        p_title.Font.Size = 18
        p_title.Font.Bold = True
        p_title.Font.Color.RGB = 6245376  # Color #004C5F
        wordDoc.Content.InsertParagraphAfter()
        
        # Buscar los bloques de celdas por palabra clave en la hoja
        for key, titulo_tabla in titleMapping.items():
            foundRange = excelSheet.Cells.Find(
                What=key, 
                LookIn=xlValues, 
                LookAt=xlPart, 
                SearchOrder=xlByRows, 
                SearchDirection=xlNext, 
                MatchCase=False
            )
            
            if foundRange is not None:
                firstAddress = foundRange.Address
                while True:
                    # Selecciona la región actual del bloque (Línea 11 de tu VBA: CurrentRegion)
                    tblRange = foundRange.CurrentRegion
                    if tblRange is not None and tblRange.Cells.Count > 1:
                        
                        try:
                            # Copia el rango como imagen fija (Línea 18 de tu VBA: CopyPicture)
                            tblRange.CopyPicture(Appearance=xlScreen, Format=xlPicture)
                            
                            # Insertar un párrafo nuevo en Word y PEGAR la imagen
                            wordDoc.Content.InsertParagraphAfter()
                            rango_pegado = wordDoc.Paragraphs.Last.Range
                            rango_pegado.Paste()
                            
                            # Añadir el rótulo superior (Caption) exactamente como tu macro
                            if wordDoc.InlineShapes.Count > 0:
                                ultima_img = wordDoc.InlineShapes(wordDoc.InlineShapes.Count)
                                ultima_img.Range.Select()
                                wordApp.Selection.InsertCaption(
                                    Label="Tabla", 
                                    Title=f": Resultados {titulo_tabla} - {excelSheet.Name}", 
                                    Position=wdCaptionPositionAbove
                                )
                                wordDoc.Content.InsertParagraphAfter()
                        except Exception:
                            pass
                            
                    # Seguir buscando si hay más bloques del mismo tipo en la hoja
                    foundRange = excelSheet.Cells.FindNext(foundRange)
                    if foundRange is None or foundRange.Address == firstAddress:
                        break
                        
    # Guardar el documento de Word en la carpeta local elegida
    archivo_final = os.path.join(output_dir, "Informe_Flujo_de_Potencia_Oficial.docx")
    wordDoc.SaveAs2(archivo_final)
    
    # Cerrar los procesos limpiamente para que no queden abiertos en el Administrador de Tareas
    wordDoc.Close(False)
    wordApp.Quit(False)
    excelWorkbook.Close(False)
    excelApp.Quit()
    
    return archivo_final

# --- 3. ACCIÓN DEL BOTÓN ---
if uploaded_file is not None:
    if st.button("🚀 Copiar Tablas como Imagen y Generar Word"):
        with st.spinner("Ejecutando copia nativa de Office en tu máquina local con Windows..."):
            try:
                # Guardar el Excel temporalmente en tu ruta local para que Excel pueda abrir el archivo físico
                temp_excel_path = os.path.join(ruta_destino, uploaded_file.name)
                with open(temp_excel_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Ejecutar el motor idéntico a tu VBA
                ruta_final_docx = ejecutar_traspaso_imagenes_vba(temp_excel_path, ruta_destino)
                
                st.success("✨ ¡Tablas copiadas como imagen fijas e informe generado con éxito!")
                st.info(f"💾 El archivo Word se guardó en tu computadora en: {ruta_final_docx}")
                
                # Limpiar el Excel temporal
                if os.path.exists(temp_excel_path):
                    os.remove(temp_excel_path)
            except Exception as e:
                st.error(f"Error al interactuar con Office. Asegúrate de cerrar Word y Excel antes de ejecutar el botón: {e}")
