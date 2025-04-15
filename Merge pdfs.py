import PyPDF2
import os
import re

def extract_number(filename):
    """Extrae el número de un nombre de archivo (si existe) para ordenarlos correctamente."""
    match = re.search(r'\d+', filename)
    return int(match.group()) if match else float('inf')  # Si no hay número, lo manda al final

def merge_pdfs(input_folder, output_pdf):
    """Fusiona PDFs en el orden numérico de su nombre."""
    pdf_files = [file for file in os.listdir(input_folder) if file.lower().endswith(".pdf")]
    pdf_files.sort(key=extract_number)  # Ordenar por número en el nombre

    merger = PyPDF2.PdfMerger()

    for pdf in pdf_files:
        pdf_path = os.path.join(input_folder, pdf)
        print(f"Añadiendo: {pdf}")  # Mensaje de depuración
        merger.append(pdf_path)  # Agregar el PDF sin alterar su contenido

    # Guardar el PDF final
    output_path = os.path.join(input_folder, output_pdf)
    merger.write(output_path)
    merger.close()
    print(f"✅ PDF fusionado con éxito: {output_path}")

# 📂 Rutas
input_folder = r"H:\Dossiers\Proletariado"  # Carpeta con los PDFs
output_pdf = "Dossier Proletariado Final.pdf"  # Nombre del PDF final

merge_pdfs(input_folder, output_pdf)
