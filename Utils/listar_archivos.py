import os
import sys
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

def listar_archivos(directorio):
    datos_archivos = []
    # Recorrer el directorio, incluyendo subdirectorios
    for raiz, _, archivos in os.walk(directorio):
        for archivo in archivos:
            ruta_completa = os.path.join(raiz, archivo)
            try:
                tamaño = os.path.getsize(ruta_completa)
                mod_time = datetime.fromtimestamp(os.path.getmtime(ruta_completa))
                cre_time = datetime.fromtimestamp(os.path.getctime(ruta_completa))
                permisos = oct(os.stat(ruta_completa).st_mode)[-3:]
                datos_archivos.append([archivo, ruta_completa, tamaño, mod_time, cre_time, permisos])
            except Exception as e:
                print(f"Error processing {ruta_completa}: {e}")
    return datos_archivos

def guardar_en_excel(datos_archivos, nombre_archivo_excel):
    wb = Workbook()
    ws = wb.active
    ws.title = "Detalles de Archivos"

    # Agregar encabezados y estilo
    headers = ['Nombre de Archivo', 'Ruta', 'Tamaño (bytes)', 'Fecha de Modificación', 'Fecha de Creación', 'Permisos']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # Agregar datos a la hoja
    for data in datos_archivos:
        nombre_archivo, ruta, tamaño, mod_time, cre_time, permisos = data
        mod_time = mod_time.strftime('%Y-%m-%d %H:%M:%S')
        cre_time = cre_time.strftime('%Y-%m-%d %H:%M:%S')
        link = f'=HYPERLINK("{ruta}", "Link")'
        ws.append([nombre_archivo, link, tamaño, mod_time, cre_time, permisos])

    # Ajustar el ancho de las columnas
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 15

    # Guardar el libro de Excel
    wb.save(nombre_archivo_excel)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        directorio = sys.argv[1]
    else:
        print("Por favor, especifica la ruta del directorio.")
        sys.exit(1)

    datos_archivos = listar_archivos(directorio)
    nombre_archivo_excel = 'Detalles_de_Archivos.xlsx'
    guardar_en_excel(datos_archivos, nombre_archivo_excel)
