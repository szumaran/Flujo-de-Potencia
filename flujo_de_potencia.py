import streamlit as st
import pandas as pd
import os
import time
import win32com.client as win32

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Plataforma Flujo de Potencia", page_icon="⚡", layout="wide")

st.title("⚡ Plataforma Flujo de Potencia")
st.write("Generación de informes mediante automatización nativa de Office (VBA en Python).")
st.markdown("---")

# Input para que el usuario ingrese la ruta local de su máquina donde se guardará el archivo final
ruta_destino = st.text_input("📁 Ingresa la ruta de la carpeta local donde deseas guardar el Word final:", value="C:\\Users\\Public\\Documents")
uploaded_file = st.file_uploader("Selecciona el archivo Excel de Resultados (EFP_RES_...)", type=["xlsx", "xlsm"])

def ejecutar_macro_pasted_tables(file_path, output_dir):
    """
    Replica de forma exacta la lógica de tu código VBA usando win32com para automatizar
    Excel y Word de forma nativa en Windows, copiando las tablas como imagen para no romper el formato.
    """
    # Constantes de Office equivalentes al código de VBA
    xlValues = -4163
    xlPart = 2
    xlByRows = 1
    xlNext = 1
    xlScreen = 1
    xlPicture = -4147
    wdCaptionPositionAbove = 0
    
    # 1. Inicializar Word
    wordApp = win32.DispatchEx("Word.Application")
    wordApp.Visible = False
    wordDoc = wordApp.Documents.Add()
    
    # Configurar rótulo de "Tabla" si no existe
    caption_exists = False
    for caption in wordApp.CaptionLabels:
        if caption.Name == "Tabla":
            caption_exists = True
            break
    if not caption_exists:
        wordApp.CaptionLabels.Add(Name="Tabla")
        
    # 2. Inicializar Excel en segundo plano
    excelApp = win32.DispatchEx("Excel.Application")
    excelApp.Visible = False
    excelWorkbook = excelApp.Workbooks.Open(file_path)
    
    # Mapeo de títulos basado en tu macro original y las columnas reales del Excel
    titleMapping = {
        "Línea\n[MVA]": "Cargabilidad Líneas (MVA)",
        "Línea\n[kA]": "Cargabilidad Líneas (kA)",
        "Transformador": "Cargabilidad Transformadores",
        "Barra": "Regulación de tensión"
    }
    
    # Recorrer todas las hojas del Excel (Escenarios)
    for excelSheet in excelWorkbook.Sheets:
        
        # Insertar título del caso (Línea 5-6 del VBA original)
        p_title = wordDoc.Content.Paragraphs.Add().Range
        p_title.Text = f"Resultados {excelSheet.Name}"
        p_title.Font.Name = "Ubuntu"
        p_title.Font.Size = 18
        p_title.Font.Bold = True
        p_title.Font.Color.RGB = 6245376 # Color azul oscuro #004C5F equivalente en entero de Office
        p_title.ParagraphFormat.Alignment = 0 # Izquierda
        wordDoc.Content.InsertParagraphAfter()
        
        # Buscar coincidencias de tablas dentro de la hoja actual
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
                    tblRange = foundRange.CurrentRegion
                    if tblRange is not None and tblRange.Cells.Count > 1:
                        
                        startCol = tblRange.Column
                        endCol = tblRange.Column + tblRange.Columns.Count - 1
                        startRow = tblRange.Row
                        endRow = tblRange.Row + tblRange.Rows.Count - 1
                        
                        # Definir el rango exacto de la tabla encontrada
                        exact_range = excelSheet.Range(excelSheet.Cells(startRow, startCol), excelSheet.Cells(endRow, endCol))
                        
                        # Copiar la tabla como imagen respetando el formato visual de Excel (Línea 18 del VBA)
                        try:
                            exact_range.CopyPicture(Appearance=xlScreen, Format=xlPicture)
                            wordDoc.Content.InsertParagraphAfter()
                            wordDoc.Paragraphs.Last.Range.Paste()
                            
                            # Configuración de rótulo superior para la tabla pegada
                            if wordDoc.InlineShapes.Count > 0:
                                img = wordDoc.InlineShapes(wordDoc.InlineShapes.Count)
                                img.Range.Select()
                                wordApp.Selection.InsertCaption(
                                    Label="Tabla", 
                                    Title=f": Resultados {titulo_tabla} - {excelSheet.Name}", 
                                    Position=wdCaptionPositionAbove
                                )
                                wordApp.Selection.ParagraphFormat.Alignment = 0
                                wordDoc.Content.InsertParagraphAfter()
                        except Exception as e:
                            pass
                            
                    # Continuar buscando el siguiente bloque en la hoja
                    foundRange = excelSheet.Cells.FindNext(foundRange)
                    if foundRange is None or foundRange.Address == firstAddress:
                        break
                        
    # Guardar el documento Word final en la ruta local definida por el usuario
    savePath = os.path.join(output_dir, "Informe_Flujo_de_Potencia_Oficial.docx")
    wordDoc.SaveAs2(savePath)
    
    # Cerrar aplicaciones limpiamente
    wordDoc.Close(False)
    wordApp.Quit(False)
    excelWorkbook.Close(False)
    excelApp.Quit()
    
    return savePath

# --- 5. ACCIÓN DEL BOTÓN ---
if uploaded_file is not None:
    if st.button("🚀 Ejecutar Copiado de Tablas Nativo"):
        with st.spinner("Procesando archivos mediante macros COM nativas de Office..."):
            try:
                # Guardar temporalmente el archivo subido en la carpeta de destino para que Excel lo pueda abrir
                temp_excel_path = os.path.join(ruta_destino, uploaded_file.name)
                with open(temp_excel_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Ejecutar el motor de traspaso COM idéntico al de VBA
                ruta_final_docx = ejecutar_macro_pasted_tables(temp_excel_path, ruta_destino)
                
                st.success(f"✨ ¡Tablas copiadas e informe generado con éxito sin deformaciones!")
                st.info(f"💾 El documento ha sido guardado de forma nativa en tu equipo en la ruta: {ruta_final_docx}")
                
                # Eliminar el archivo temporal de Excel
                if os.path.exists(temp_excel_path):
                    os.remove(temp_excel_path)
            except Exception as e:
                st.error(f"Ocurrió un error al interactuar con las aplicaciones de Office: {e}")
