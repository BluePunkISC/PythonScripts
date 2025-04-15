import fitz  # PyMuPDF
from PIL import Image
import os

def compress_image_to_1mb(image_path, output_path):
    try:
        img = Image.open(image_path)
        quality = 95  # Calidad inicial
        while True:
            img.save(output_path, "JPEG", quality=quality)
            size = os.path.getsize(output_path) / (1024 * 1024)  # Tamaño en MB
            if size <= 1 or quality <= 10:
                break
            quality -= 5
        print(f"Imagen comprimida a {size:.2f} MB con calidad {quality}.")
    except Exception as e:
        print(f"Error al procesar {image_path}: {e}")

def pdf_to_images(pdf_path, output_dir):
    # Crear carpeta de salida si no existe
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Abrir el archivo PDF
    pdf_document = fitz.open(pdf_path)
    for page_number in range(len(pdf_document)):
        # Extraer página como imagen
        page = pdf_document.load_page(page_number)
        pix = page.get_pixmap(dpi=150)  # Resolución ajustable
        image_path = os.path.join(output_dir, f"page_{page_number + 1}.png")

        # Guardar imagen temporal
        pix.save(image_path)

        # Comprimir la imagen a menos de 1 MB
        compressed_image_path = os.path.join(output_dir, f"page_{page_number + 1}_compressed.jpg")
        compress_image_to_1mb(image_path, compressed_image_path)

        # Eliminar imagen temporal
        os.remove(image_path)

    pdf_document.close()
    print("Conversión completada.")

# Rutas de entrada y salida
pdf_path = r"H:\Dossier\Rider A.pdf"  # Ruta del PDF
output_dir = r"H:\Dossier"  # Carpeta donde se guardarán las imágenes

# Llamada a la función
pdf_to_images(pdf_path, output_dir)
