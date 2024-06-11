import openpyxl

file_path = 'C:/Users/lmonroy/Tema/MASTER ORDENES.xlsm'
repaired_file_path = 'C:/Users/lmonroy/Tema/MASTER_ORDENES_REPARADO.xlsm'

try:
    # Load the workbook
    workbook = openpyxl.load_workbook(file_path)

    # Save the workbook to a new file
    workbook.save(repaired_file_path)
    repair_status = "Archivo reparado y guardado como 'MASTER_ORDENES_REPARADO.xlsm'."
except Exception as e:
    repair_status = str(e)

repair_status
