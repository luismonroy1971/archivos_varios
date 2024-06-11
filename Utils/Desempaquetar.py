import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import zipfile
import os

def select_and_extract_zip():
    # Crear una ventana oculta para la selección del archivo
    root = tk.Tk()
    root.withdraw()
    
    # Abrir un cuadro de diálogo para seleccionar el archivo zip
    zip_path = filedialog.askopenfilename(
        title="Selecciona un archivo ZIP",
        filetypes=[("Archivos ZIP", "*.zip")]
    )
    
    if zip_path:
        # Seleccionar la ubicación de extracción
        extract_to = filedialog.askdirectory(
            title="Selecciona la carpeta de destino"
        )
        
        if extract_to:
            try:
                # Extraer el archivo zip
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    for member in zip_ref.namelist():
                        member_path = os.path.join(extract_to, member)
                        if not os.path.abspath(member_path).startswith(os.path.abspath(extract_to)):
                            raise Exception("Tentativa de extracción insegura detectada!")
                        zip_ref.extract(member, extract_to)
                messagebox.showinfo("Éxito", "El archivo se ha desempaquetado correctamente.")
            except zipfile.BadZipFile:
                messagebox.showerror("Error", "El archivo seleccionado no es un archivo zip válido.")
            except Exception as e:
                messagebox.showerror("Error", f"Ha ocurrido un error: {str(e)}")
        else:
            messagebox.showwarning("Cancelado", "La operación ha sido cancelada.")
    else:
        messagebox.showwarning("Cancelado", "La operación ha sido cancelada.")

# Ejecutar la función principal
if __name__ == "__main__":
    select_and_extract_zip()
