import os
import argparse
from PyPDF2 import PdfMerger, PdfReader

def es_pdf_valido(ruta):
    try:
        with open(ruta, "rb") as archivo:
            PdfReader(archivo)
        return True
    except:
        return False

def combinar_pdfs(ruta_principal, ruta_destino):
    if not os.path.exists(ruta_destino):
        os.makedirs(ruta_destino)

    for raiz, subcarpetas, archivos in os.walk(ruta_principal):
        if raiz == ruta_destino:
            continue
        pdf_merger = PdfMerger()
        pdf_files = [archivo for archivo in archivos if archivo.endswith('.pdf')]

        for archivo in pdf_files:
            ruta_completa = os.path.join(raiz, archivo)
            if es_pdf_valido(ruta_completa):
                try:
                    pdf_merger.append(ruta_completa)
                except Exception as e:
                    print(f"No se pudo añadir {ruta_completa}: {e}")

        if pdf_merger.inputs:
            nombre_combinado = f"Combinado_{os.path.basename(raiz)}.pdf"
            ruta_combinado = os.path.join(ruta_destino, nombre_combinado)
            pdf_merger.write(ruta_combinado)
            pdf_merger.close()
            print(f"PDF combinado creado en: {ruta_combinado}")
        else:
            pdf_merger.close()

def main():
    parser = argparse.ArgumentParser(description='Combinar PDFs desde subcarpetas.')
    parser.add_argument('ruta_principal', type=str, help='Ruta principal donde están las subcarpetas con los PDFs')
    parser.add_argument('ruta_destino', type=str, help='Ruta donde se guardarán los PDFs combinados')

    args = parser.parse_args()
    combinar_pdfs(args.ruta_principal, args.ruta_destino)

if __name__ == '__main__':
    main()

print("Proceso completado.")
