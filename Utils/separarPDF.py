import os
from PyPDF2 import PdfReader, PdfWriter

def split_pdf(input_pdf_path, output_dir):
    # Asegurarse de que la ruta de salida exista
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Abrir el archivo PDF
    with open(input_pdf_path, 'rb') as pdf_file:
        pdf_reader = PdfReader(pdf_file)
        num_pages = len(pdf_reader.pages)

        # Iterar sobre cada página y guardar cada una como un nuevo archivo PDF
        for page_num in range(num_pages):
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_num])

            output_filename = os.path.join(output_dir, f'page_{page_num + 1}.pdf')
            with open(output_filename, 'wb') as output_pdf_file:
                pdf_writer.write(output_pdf_file)

            print(f'Guardado: {output_filename}')

if __name__ == "__main__":
    input_pdf_path = input("Ingrese la ruta completa del archivo PDF de origen: ")
    output_dir = input("Ingrese la ruta completa de la carpeta de destino: ")
    
    split_pdf(input_pdf_path, output_dir)
    print("Proceso completado.")
