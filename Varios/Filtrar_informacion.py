import pandas as pd

def procesar_datos(ruta_archivo):
    # Cargar el archivo Excel
    df = pd.read_excel(ruta_archivo)
    
    # Asumiendo que el nombre del proyecto puede ser extraído de "RUTA COMPLETA"
    # Por ejemplo, si "RUTA COMPLETA" contiene el nombre del proyecto al inicio seguido de un delimitador como "/"
    df['PROYECTO'] = df['RUTA COMPLETA'].str.split('/').str[0]

    # Lista de eventos para contar
    eventos = ['charlas', 'capacitaciones', 'ATS', 'RACSI', 'atenciones médicas']
    
    # Diccionario para almacenar los resultados
    resultados = {}
    
    # Agrupar por proyecto y contar cada tipo de evento
    for evento in eventos:
        df[evento] = df['RUTA COMPLETA'].str.contains(evento, case=False, na=False)
        resultados[evento] = df.groupby('PROYECTO')[evento].sum()
    
    return resultados

resultados = procesar_datos('C:/Users/lmonroy/Tema/archivos_red/Origen/FILTRADOS_PROYECTOS/LD PROYECTOS TEMA GENERAL.xlsx')
print(resultados)
