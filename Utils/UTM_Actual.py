from opencage.geocoder import OpenCageGeocode
import utm
import requests

def verify_api_key(api_key):
    """
    Verifica si la clave de API de OpenCage es válida.
    
    :param api_key: Clave de API de OpenCage
    :return: Booleano indicando si la clave es válida
    """
    url = f"https://api.opencagedata.com/geocode/v1/json?q=51.952659,7.632473&key={api_key}"
    response = requests.get(url)
    return response.status_code == 200

def get_geographic_coordinates(address, api_key):
    """
    Obtiene las coordenadas geográficas (latitud y longitud) de una dirección utilizando OpenCage API.
    
    :param address: Dirección o ubicación para geocodificar
    :param api_key: Clave de API de OpenCage
    :return: Tupla (latitud, longitud)
    """
    geocoder = OpenCageGeocode(api_key)
    result = geocoder.geocode(address)
    
    if result and len(result):
        return result[0]['geometry']['lat'], result[0]['geometry']['lng']
    else:
        raise ValueError("No se pudieron obtener las coordenadas geográficas")

def latlon_to_utm(lat, lon):
    """
    Convierte coordenadas geográficas a coordenadas UTM.
    
    :param lat: Latitud
    :param lon: Longitud
    :return: Tupla (easting, northing, zone_number, zone_letter)
    """
    return utm.from_latlon(lat, lon)

if __name__ == "__main__":
    # Reemplaza con tu dirección y clave de API
    address = "Your Address or Location"
    api_key = "YOUR_OPENCAGE_API_KEY"

    # Verifica la clave de API
    if not verify_api_key(api_key):
        print("Error: Your API key is not authorized. You may have entered it incorrectly.")
    else:
        try:
            # Obtiene las coordenadas geográficas
            latitude, longitude = get_geographic_coordinates(address, api_key)
            print(f"Latitud: {latitude}, Longitud: {longitude}")
            
            # Convierte a coordenadas UTM
            easting, northing, zone_number, zone_letter = latlon_to_utm(latitude, longitude)
            print(f"Coordenadas UTM: Easting: {easting}, Northing: {northing}, Zona: {zone_number}{zone_letter}")
        except Exception as e:
            print(f"Error: {e}")
