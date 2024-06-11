import os
import openpyxl
from openpyxl.styles import Font
from tkinter import Tk, filedialog

def explorar_carpeta_principal(carpeta_principal):
    estructura = []
    for root, dirs, files in os.walk(carpeta_principal):
        nivel = root.replace(carpeta_principal, '').count(os.sep)
        carpeta_actual = os.path.basename(root) if root != carpeta_principal else root
        if not files:  # Si no hay archivos en la carpeta, no se agrega nada
            continue
        for f in files:
            niveles = root.replace(carpeta_principal, '').split(os.sep)
            niveles = [carpeta_principal] + niveles + [f]
            estructura.append(niveles)
    return estructura

def guardar_estructura_en_excel(estructura, archivo_excel):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Estructura de Carpetas"
    
    # Establecer el encabezado
    max_niveles = max(len(niveles) for niveles in estructura)
    encabezados = ["Nivel " + str(i) for i in range(max_niveles)] + ["Archivo Completo"]
    ws.append(encabezados)
    
    for cell in ws["1:1"]:
        cell.font = Font(bold=True)
    
    # Añadir la estructura al archivo Excel
    for niveles in estructura:
        fila = niveles + [''] * (max_niveles - len(niveles))
        ruta_completa = os.path.join(*niveles)
        fila.append(ruta_completa)
        ws.append(fila)
    
    wb.save(archivo_excel)

def seleccionar_carpeta(titulo):
    root = Tk()
    root.withdraw()  # Ocultar la ventana principal de Tkinter
    carpeta = filedialog.askdirectory(title=titulo)
    root.destroy()
    return carpeta

def seleccionar_archivo(titulo):
    root = Tk()
    root.withdraw()  # Ocultar la ventana principal de Tkinter
    archivo = filedialog.asksaveasfilename(title=titulo, defaultextension=".xlsx",
                                           filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    return archivo

def main():
    carpeta_principal = seleccionar_carpeta("Seleccione la carpeta principal a explorar")
    archivo_excel = seleccionar_archivo("Seleccione el archivo Excel donde se grabará la estructura")
    
    if carpeta_principal and archivo_excel:
        estructura = explorar_carpeta_principal(carpeta_principal)
        guardar_estructura_en_excel(estructura, archivo_excel)
        print(f"Estructura guardada en {archivo_excel}")
    else:
        print("No se seleccionó una carpeta o archivo válido.")

# Descomenta la siguiente línea para ejecutar el programa
main()
