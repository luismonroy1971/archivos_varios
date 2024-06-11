import pandas as pd
import math

# Constantes para la conversión UTM a Lat/Lon
K0 = 0.9996
E = 0.00669438
E2 = E * E
E3 = E2 * E
E_P2 = E / (1.0 - E)

SQRT_E = math.sqrt(1 - E)
_E = (1 - SQRT_E) / (1 + SQRT_E)
_E2 = _E * _E
_E3 = _E2 * _E
_E4 = _E3 * _E
_E5 = _E4 * _E

M1 = 1 - E / 4.0 - 3 * E2 / 64.0 - 5 * E3 / 256.0
M2 = 3 * E / 8.0 + 3 * E2 / 32.0 + 45 * E3 / 1024.0
M3 = 15 * E2 / 256.0 + 45 * E3 / 1024.0
M4 = 35 * E3 / 3072.0

R = 6378137

def utm_to_latlon(easting, northing, zone_number, northern_hemisphere=True):
    x = easting - 500000.0
    y = northing
    if not northern_hemisphere:
        y -= 10000000.0

    m = y / K0
    mu = m / (R * M1)

    p_rad = (mu +
             M2 * math.sin(2 * mu) +
             M3 * math.sin(4 * mu) +
             M4 * math.sin(6 * mu))

    p_sin = math.sin(p_rad)
    p_sin2 = p_sin * p_sin

    p_cos = math.cos(p_rad)

    p_tan = p_sin / p_cos
    p_tan2 = p_tan * p_tan
    p_tan4 = p_tan2 * p_tan2

    ep_sin = 1 - E * p_sin2
    ep_sin_sqrt = math.sqrt(1 - E * p_sin2)

    n = R / ep_sin_sqrt
    r = (1 - E) / ep_sin

    c = _E * p_cos * p_cos
    c2 = c * c

    d = x / (n * K0)
    d2 = d * d
    d3 = d2 * d
    d4 = d3 * d
    d5 = d4 * d
    d6 = d5 * d

    latitude = (p_rad - (p_tan / r) *
                (d2 / 2 -
                 d4 / 24 * (5 + 3 * p_tan2 + 10 * c - 4 * c2 - 9 * E_P2)) +
                d6 / 720 * (61 + 90 * p_tan2 + 298 * c + 45 * p_tan4 - 252 * E_P2 - 3 * c2))
    longitude = (d -
                 d3 / 6 * (1 + 2 * p_tan2 + c) +
                 d5 / 120 * (5 - 2 * c + 28 * p_tan2 - 3 * c2 + 8 * E_P2 + 24 * p_tan4)) / p_cos

    latitude = math.degrees(latitude)
    longitude = math.degrees(longitude) + zone_number * 6 - 183

    return latitude, longitude

# Load the CSV file with the correct delimiter
file_path = '/Users/lmonroy/Tema/UTM_LATLON/COORDENADAS CSV.csv'
df = pd.read_csv(file_path, delimiter=';')

# Apply the UTM to lat/lon conversion
df['Latitud'], df['Longitud'] = zip(*df.apply(lambda row: utm_to_latlon(row['E'], row['N'], row['ZONA'], True), axis=1))

# Save the results to a new CSV file
output_path = file_path.replace('.csv', '_latlon.csv')
df.to_csv(output_path, index=False)

output_path
