import os
import sys
import pandas as pd

def read_excel_files(folder_path, search_string):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    summary_data = []
    
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, sheet_name='Resumen')
        
        # Filtrar las filas que contienen la cadena de búsqueda en cualquier columna, ignorando mayúsculas/minúsculas
        df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(search_string, case=False).any(), axis=1)]
        
        # Extraer datos específicos de las columnas
        df_filtered['ÁREA'] = df_filtered.iloc[:, 0].str[3:6]
        df_filtered['PERIODO'] = df_filtered.iloc[:, 0].str[7:11]
        df_filtered['PROYECTO'] = df_filtered.iloc[:, 2]  # Columna C después de ignorar la columna B
        
        # Agregar los datos filtrados y extraídos a la lista de datos resumen
        summary_data.append(df_filtered[['ÁREA', 'PERIODO', 'PROYECTO']])
    
    # Concatenar todos los datos en un DataFrame único
    summary_df = pd.concat(summary_data, ignore_index=True)
    
    with pd.ExcelWriter('RESULTADOS.xlsx', engine='xlsxwriter') as writer:
        # Escribir la hoja Resumen General
        summary_df.to_excel(writer, sheet_name='Resumen General', index=False)
        
        # Crear la hoja Tabla Final
        final_table = summary_df.groupby('PROYECTO').size().reset_index(name='Cantidad')
        final_table.to_excel(writer, sheet_name='Tabla Final', index=False)
        
        # Agregar título a la hoja Tabla Final
        workbook = writer.book
        worksheet = writer.sheets['Tabla Final']
        worksheet.write('A1', f'Resumen de {search_string}')
    
    return 'RESULTADOS.xlsx creado con éxito'

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Uso: python read_excel_files.py <ruta_de_carpeta> <cadena_a_buscar>")
    else:
        folder_path = sys.argv[1]
        search_string = sys.argv[2]
        result = read_excel_files(folder_path, search_string)
        print(result)

# python Resumen_Doc.py "C:/Users/lmonroy/Tema/archivos_red/Origen/FILTRADOS_PROYECTOS/" "Charla Diaria"