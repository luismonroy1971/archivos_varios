import tkinter as tk
from tkinter import filedialog
import pandas as pd
import json

def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal de Tkinter
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
    )
    return archivo

def excel_a_json(archivo_excel):
    # Leer el archivo Excel asegurándose de que los campos DNINumber y DNIType se lean como texto
    df = pd.read_excel(archivo_excel, dtype={'DNINumber': str, 'DNIType': str})
    
    # Asegurar que los campos Name1, Name2, y Name3 tengan valores vacíos si están vacíos
    for name_field in ['Name1', 'Name2', 'Name3']:
        if name_field in df.columns:
            df[name_field].fillna("", inplace=True)
    
    # Convierte el DataFrame a una lista de diccionarios (JSON)
    json_data = df.to_dict(orient='records')
    
    return json_data

def guardar_json(data, archivo_salida):
    with open(archivo_salida, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def main():
    archivo_excel = seleccionar_archivo()
    if archivo_excel:
        json_data = excel_a_json(archivo_excel)
        guardar_json(json_data, "salida.json")
        print("El archivo JSON ha sido creado exitosamente y guardado como salida.json")

if __name__ == "__main__":
    main()
