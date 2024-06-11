import os
import sys
import pandas as pd

def read_excel_files(folder_path, search_string):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    summary_data = []
    
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, sheet_name='Resumen')
        
        df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(search_string).any(), axis=1)].copy()
        df_filtered.loc[:, 'ÁREA'] = df_filtered.iloc[:, 0].str[3:6]
        df_filtered.loc[:, 'PERIODO'] = df_filtered.iloc[:, 0].str[7:11]
        df_filtered.loc[:, 'PROYECTO'] = df_filtered.iloc[:, 1]
        
        summary_data.append(df_filtered[['ÁREA', 'PERIODO', 'PROYECTO']])
    
    summary_df = pd.concat(summary_data, ignore_index=True)
    
    # Usando xlsxwriter como el motor para pandas ExcelWriter
    with pd.ExcelWriter('RESULTADOS.xlsx', engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, sheet_name='Resumen General', index=False)
        final_table = summary_df.groupby(['PROYECTO', 'PERIODO']).size().reset_index(name='Cantidad')
        final_table.to_excel(writer, sheet_name='Tabla Final', index=False)

        # Accediendo al objeto workbook y worksheet
        workbook  = writer.book
        worksheet = writer.sheets['Tabla Final']
        
        # Agregando título usando xlsxwriter
        worksheet.write('A1', f'Resumen de {search_string}')
    
    return 'RESULTADOS.xlsx created successfully'

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python read_excel_files.py <folder_path> <search_string>")
    else:
        folder_path = sys.argv[1]
        search_string = sys.argv[2]
        result = read_excel_files(folder_path, search_string)
        print(result)
