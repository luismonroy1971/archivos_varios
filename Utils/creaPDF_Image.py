import os
from PIL import Image
from fpdf import FPDF
from tkinter import Tk
from tkinter.filedialog import askdirectory

def crear_pdf_con_imagenes(carpeta):
    # Verificar que la carpeta existe
    if not os.path.isdir(carpeta):
        print("La carpeta no existe.")
        return
    
    # Obtener lista de archivos en la carpeta
    archivos = os.listdir(carpeta)
    
    # Filtrar archivos de imagen
    imagenes = [archivo for archivo in archivos if archivo.lower().endswith(('png', 'jpg', 'jpeg', 'bmp', 'gif'))]
    
    if not imagenes:
        print("No se encontraron imágenes en la carpeta.")
        return
    
    # Crear objeto PDF
    pdf = FPDF()
    
    # Definir dimensiones de la página A4 en mm (1 cm = 10 mm)
    page_width_mm, page_height_mm = 210, 297
    max_width_mm = 150  # 15 cm = 150 mm
    space_between_images_mm = 50  # 5 cm = 50 mm
    
    y_offset_mm = 0

    # Añadir la primera página
    pdf.add_page()
    
    for imagen in imagenes:
        ruta_imagen = os.path.join(carpeta, imagen)
        
        # Abrir la imagen
        img = Image.open(ruta_imagen)
        
        # Convertir imagen a RGB si es necesario
        if img.mode in ("RGBA", "P"): 
            img = img.convert("RGB")
        
        # Guardar la imagen en un archivo temporal
        ruta_temp = os.path.join(carpeta, "temp.jpg")
        img.save(ruta_temp, "JPEG")
        
        # Obtener dimensiones de la imagen
        img_width, img_height = img.size
        img_aspect_ratio = img_width / img_height
        
        # Calcular el tamaño de la imagen para mantener la proporción y ajustar por el ancho
        new_width_mm = max_width_mm
        new_height_mm = new_width_mm / img_aspect_ratio
        
        x_offset_mm = (page_width_mm - new_width_mm) / 2
        
        # Verificar si necesitamos una nueva página para la imagen actual
        if y_offset_mm + new_height_mm > page_height_mm:
            pdf.add_page()
            y_offset_mm = 0
        
        # Agregar la imagen al PDF centrada horizontalmente
        pdf.image(ruta_temp, x_offset_mm, y_offset_mm, new_width_mm, new_height_mm)
        
        # Incrementar el offset vertical para la próxima imagen
        y_offset_mm += new_height_mm + space_between_images_mm
        
        # Eliminar la imagen temporal
        os.remove(ruta_temp)
    
    # Guardar el PDF
    ruta_pdf = os.path.join(carpeta, "imagenes.pdf")
    pdf.output(ruta_pdf)
    print(f"PDF creado exitosamente: {ruta_pdf}")

if __name__ == "__main__":
    # Inicializar Tkinter y ocultar la ventana principal
    root = Tk()
    root.withdraw()
    
    # Abrir el cuadro de diálogo para seleccionar la carpeta
    carpeta = askdirectory(title="Selecciona la carpeta que contiene las imágenes")
    
    if carpeta:
        crear_pdf_con_imagenes(carpeta)
    else:
        print("No se seleccionó ninguna carpeta.")
