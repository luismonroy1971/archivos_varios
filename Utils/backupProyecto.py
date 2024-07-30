import pandas as pd
import os
import shutil
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

def seleccionar_destino():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    carpeta_destino = filedialog.askdirectory(title="Selecciona el destino para el backup")
    return carpeta_destino

def crear_estructura_carpetas(row, base_path):
    codigo_proyecto = str(row['CÓDIGO PROYECTO'])
    empresa = str(row['EMPRESA'])
    doc_identidad = str(row['DOC IDENTIDAD'])
    link = row['LINK']
    
    # Crear las carpetas según la estructura
    path_proyecto = os.path.join(base_path, codigo_proyecto)
    path_personal = os.path.join(path_proyecto, 'PERSONAL')
    path_empresa = os.path.join(path_personal, empresa)
    path_doc_identidad = os.path.join(path_empresa, doc_identidad)
    
    os.makedirs(path_doc_identidad, exist_ok=True)
    
    # Copiar los archivos desde el hipervínculo
    if os.path.exists(link):
        shutil.copy(link, path_doc_identidad)
    else:
        print(f"No se encontró el archivo: {link}")

def obtener_ruta_documentos():
    from pathlib import Path
    return str(Path.home() / "Documents")

def main():
    # Obtener la ruta a la carpeta "Documentos" del usuario
    carpeta_documentos = obtener_ruta_documentos()
    
    # Ruta al archivo de Excel dentro de la carpeta "Documentos"
    archivo_excel = os.path.join(carpeta_documentos, 'App Habilitaciones 1.0.xlsm')
    
    # Comprobar si el archivo existe
    if not os.path.exists(archivo_excel):
        print(f"No se encontró el archivo: {archivo_excel}")
        return
    
    # Leer el archivo de Excel
    hoja_excel = 'DASHBOARD PERSONAL'
    df = pd.read_excel(archivo_excel, sheet_name=hoja_excel)
    
    # Seleccionar destino para el backup
    carpeta_destino = seleccionar_destino()
    if not carpeta_destino:
        print("No se seleccionó ningún destino. Saliendo del programa.")
        return
    
    # Iterar sobre cada fila de la tabla y crear la estructura de carpetas
    for index, row in df.iterrows():
        crear_estructura_carpetas(row, carpeta_destino)

if __name__ == "__main__":
    main()
