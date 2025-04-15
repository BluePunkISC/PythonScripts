from PIL import Image, ImageOps
import os
import re

# 🏷️ Tamaño en píxeles para A4 a 300 DPI
A4_SIZE = (2480, 3508)  # (ancho, alto)

def extract_number(filename):
    """Extrae el número de un nombre de archivo (si existe) para ordenar correctamente."""
    match = re.search(r'\d+', filename)
    return int(match.group()) if match else float('inf')  # Si no hay número, lo manda al final

def compress_image(image_path, quality=50):
    """Comprime la imagen y la ajusta al tamaño A4 sin deformarla."""
    img = Image.open(image_path)
    img = img.convert("RGB")  # Asegurar compatibilidad con PDF

    # Redimensionar manteniendo relación de aspecto
    img.thumbnail(A4_SIZE, Image.LANCZOS)

    # Crear lienzo A4 y centrar la imagen
    a4_canvas = Image.new("RGB", A4_SIZE, "white")  # Fondo blanco
    x_offset = (A4_SIZE[0] - img.width) // 2
    y_offset = (A4_SIZE[1] - img.height) // 2
    a4_canvas.paste(img, (x_offset, y_offset))

    # Guardar la imagen comprimida
    compressed_path = os.path.splitext(image_path)[0] + "_compressed.jpg"
    a4_canvas.save(compressed_path, "JPEG", quality=quality)
    
    return compressed_path

def images_to_pdf(input_folder, output_pdf):
    """Convierte imágenes numeradas en un PDF en tamaño A4 y en el orden correcto."""
    images = []
    
    # Obtener lista de imágenes ordenadas por número
    image_files = [file for file in os.listdir(input_folder) if file.lower().endswith((".jpg", ".jpeg", ".png"))]
    image_files.sort(key=extract_number)  # Orden numérico

    for file in image_files:
        image_path = os.path.join(input_folder, file)
        compressed_path = compress_image(image_path)  # Comprime y ajusta al tamaño A4
        img = Image.open(compressed_path)
        images.append(img)

    if images:
        pdf_path = os.path.join(input_folder, output_pdf)
        images[0].save(pdf_path, save_all=True, append_images=images[1:])  # Guarda como PDF
        print(f"✅ PDF creado con éxito: {pdf_path}")

        # 🔥 Eliminar imágenes comprimidas para ahorrar espacio
        for img in images:
            os.remove(img.filename)

    else:
        print("⚠️ No se encontraron imágenes en la carpeta.")

# 📂 Rutas
input_folder = r"H:\Dossiers\Proletariado"  # Carpeta con imágenes
output_pdf = "Trayectoria.pdf"  # Nombre del PDF de salida

images_to_pdf(input_folder, output_pdf)
