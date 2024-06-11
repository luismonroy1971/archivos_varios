import requests

# Configuración del token de autorización y encabezados
API_TOKEN = '2a66dec33b75527b56754cd153bafc18f7bc1c4a2aa04758eb45069202df830d'
HEADERS = {
    'Authorization': f'Bearer {API_TOKEN}',
    'Content-Type': 'application/json'
}

# Definición de la URL del endpoint
BASE_URL = 'https://api.safetyculture.io'
INSPECTIONS_URL = f'{BASE_URL}/audits/search'

# Función para obtener datos de inspecciones
def get_inspections():
    try:
        response = requests.get(INSPECTIONS_URL, headers=HEADERS)
        response.raise_for_status()  # Lanza una excepción si la solicitud falla
        return response.json()  # Devolver toda la respuesta JSON
    except requests.exceptions.HTTPError as err:
        print(f'HTTP error occurred: {err}')  # Manejo de errores HTTP
        return None
    except Exception as err:
        print(f'Other error occurred: {err}')  # Otros errores
        return None

# Obtener y verificar la estructura de los datos de inspecciones
inspections_data = get_inspections()

if inspections_data:
    print(inspections_data)
else:
    print("No se pudieron obtener datos de inspecciones.")
