import pandas as pd
import json

# Ajusta la ruta del archivo y los nombres de las columnas según sea necesario
file_path = 'C:/Users/lmonroy/Tema/MasterDB/INGRESOS DE TRABAJADORES.xlsx'
sheet_name = 'MASTERDB'
dtype_spec = {'DNINumber': str, 'DNIType': str}

# Leer la hoja de Excel
df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=dtype_spec)

# Convertir la columna 'BirthDate' a datetime, asumiendo el formato 'YYYY-MM-DD'
df['BirthDate'] = pd.to_datetime(df['BirthDate'], format='%Y-%m-%d', errors='coerce')

# Formatear las fechas al formato ISO 8601 después de asegurar que la interpretación inicial sea correcta
df['BirthDate'] = df['BirthDate'].dt.strftime('%Y-%m-%dT00:00:00Z').where(df['BirthDate'].notnull(), None)

# Llenar los campos vacíos en 'Name2' y 'Name3' con una cadena vacía
df['Name2'] = df['Name2'].fillna("")
df['Name3'] = df['Name3'].fillna("")

# Convertir cada fila en un objeto JSON y agregarlo a una lista
json_list = df.to_dict(orient='records')

# Guardar la lista de objetos JSON en un archivo
output_file_path = 'C:/trabajos/python/archivos/salida.json'
with open(output_file_path, 'w', encoding='utf-8') as file:
    json.dump(json_list, file, ensure_ascii=False, indent=4)

print("La conversión a JSON ha sido completada y guardada en:", output_file_path)
