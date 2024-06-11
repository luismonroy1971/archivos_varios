import utm

def latlon_to_utm(lat, lon):
    """
    Convierte coordenadas geográficas a coordenadas UTM.
    
    :param lat: Latitud
    :param lon: Longitud
    :return: Tupla (easting, northing, zone_number, zone_letter)
    """
    return utm.from_latlon(lat, lon)

if __name__ == "__main__":
    # Coordenadas geográficas de Lima, Perú
    latitude = -12.0464
    longitude = -77.0428
    
    # Convierte a coordenadas UTM
    easting, northing, zone_number, zone_letter = latlon_to_utm(latitude, longitude)
    print(f"Coordenadas UTM de Lima: Easting: {easting}, Northing: {northing}, Zona: {zone_number}{zone_letter}")
