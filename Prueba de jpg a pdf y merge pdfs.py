from PIL import Image
import os
import re
import pypdf  # Mejor alternativa a PyPDF2

# 🏷️ Tamaño en píxeles para A4 a 300 DPI
A4_SIZE = (2480, 3508)

def extract_number(filename):
    """Extrae el número de un nombre de archivo para ordenarlos correctamente."""
    match = re.search(r'\d+', filename)
    return int(match.group()) if match else float('inf')  # Si no hay número, lo manda al final

def compress_image(image_path, quality=50):
    """Comprime la imagen y la ajusta al tamaño A4 sin deformarla."""
    img = Image.open(image_path).convert("RGB")
    img.thumbnail(A4_SIZE, Image.LANCZOS)
    a4_canvas = Image.new("RGB", A4_SIZE, "white")
    x_offset = (A4_SIZE[0] - img.width) // 2
    y_offset = (A4_SIZE[1] - img.height) // 2
    a4_canvas.paste(img, (x_offset, y_offset))
    compressed_path = os.path.splitext(image_path)[0] + "_compressed.jpg"
    a4_canvas.save(compressed_path, "JPEG", quality=quality)
    return compressed_path

def images_to_pdf(input_folder, output_pdf):
    """Convierte imágenes numeradas en un PDF en tamaño A4 y en el orden correcto."""
    images = []
    image_files = sorted(
        [f for f in os.listdir(input_folder) if f.lower().endswith((".jpg", ".jpeg", ".png"))],
        key=extract_number
    )
    for file in image_files:
        image_path = os.path.join(input_folder, file)
        compressed_path = compress_image(image_path)
        img = Image.open(compressed_path)
        images.append(img)
    if images:
        pdf_path = os.path.join(input_folder, output_pdf)
        images[0].save(pdf_path, save_all=True, append_images=images[1:])
        print(f"✅ PDF creado con éxito: {pdf_path}")

def merge_pdfs(input_folder, output_pdf):
    """Fusiona PDFs en el orden numérico sin perder imágenes o enlaces."""
    pdf_files = sorted(
        [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")],
        key=extract_number
    )
    merger = pypdf.PdfMerger()
    for pdf in pdf_files:
        pdf_path = os.path.join(input_folder, pdf)
        print(f"📄 Añadiendo: {pdf}")
        merger.append(pdf_path)
    output_path = os.path.join(input_folder, output_pdf)
    merger.write(output_path)
    merger.close()
    print(f"✅ PDF fusionado con éxito: {output_path}")

# 📂 Rutas
input_folder = r"H:\\Dossiers\\Proletariado"
output_pdf_images = "Dossier_Proletariado_A.pdf"
output_pdf_final = "Dossier_Proletariado_Final.pdf"

# 🔹 Generar PDF de imágenes y luego fusionarlo
images_to_pdf(input_folder, output_pdf_images)
merge_pdfs(input_folder, output_pdf_final)
