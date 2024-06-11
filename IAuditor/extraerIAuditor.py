import requests
from openpyxl import Workbook

# Configuración del token de autorización y encabezados
API_TOKEN = 'tu_token_aqui'
HEADERS = {
    'Authorization': f'Bearer 2a66dec33b75527b56754cd153bafc18f7bc1c4a2aa04758eb45069202df830d',
    'Content-Type': 'application/json'
}

# Definición de URLs del API
BASE_URL = 'https://api.safetyculture.io'
ENDPOINTS = {
    'inspections': f'{BASE_URL}/audits/search',
    'users': f'{BASE_URL}/org/users',
    'templates': f'{BASE_URL}/templates/search',
    'sites': f'{BASE_URL}/sites'
}

# Función para obtener datos de un endpoint específico
def get_data(endpoint):
    try:
        response = requests.get(endpoint, headers=HEADERS)
        response.raise_for_status()  # Lanza una excepción si la solicitud falla
        return response.json() if response.text else []
    except requests.exceptions.HTTPError as err:
        print(f'HTTP error occurred: {err}')  # Manejo de errores HTTP
        return []
    except Exception as err:
        print(f'Other error occurred: {err}')  # Otros errores
        return []

# Función para agregar datos a una hoja de Excel
def add_data_to_sheet(sheet, data, headers):
    sheet.append(headers)  # Agrega los encabezados
    for item in data:
        if isinstance(item, dict):
            sheet.append([item.get(header, '') for header in headers])
        else:
            print(f'Unexpected data format: {item}')

# Crear el libro de Excel y agregar hojas
wb = Workbook()
sheets_info = {
    'Inspections': ('inspections', ['audit_id', 'name', 'created_at', 'modified_at']),
    'Users': ('users', ['id', 'firstname', 'lastname', 'email']),
    'Templates': ('templates', ['id', 'name', 'modified_at']),
    'Sites': ('sites', ['id', 'name'])
}

for sheet_name, (endpoint_key, headers) in sheets_info.items():
    sheet = wb.create_sheet(title=sheet_name)
    data = get_data(ENDPOINTS[endpoint_key])
    if 'data' in data:
        add_data_to_sheet(sheet, data['data'], headers)
    elif 'audits' in data:
        add_data_to_sheet(sheet, data['audits'], headers)
    else:
        add_data_to_sheet(sheet, data, headers)

# Eliminar la hoja por defecto que se crea con openpyxl
default_sheet = wb['Sheet']
wb.remove(default_sheet)

# Guardar el archivo Excel
wb.save('iAuditor_Data.xlsx')

print("Archivo Excel creado exitosamente con los datos de iAuditor.")
