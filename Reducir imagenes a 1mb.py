from PIL import Image
import os
def compress_image_to_1mb(input_path, output_path):
    try:
        #Abrir imagen
        img=Image.open(input_path)

        #Reducir la calidad iterativamente hasta que el tamaño sea menor o igual a 1MB
        quality=95 #Comenzar con alta calidad
        while True:
            img.save(output_path, "JPEG",quality=quality)
            size=os.path.getsize(output_path)/(1024*1024) #Convertir a MB
            if size <= 1 or quality <= 10: #Limitar la calidad minima a 10
                break
            quality -= 5 #Reducir la calidad gradualmente

        print(f"Imágen comprimida a {size:.2f}MB con calidad {quality}.")
    except Exception as e:
        print(f"Error al procesar {input_path}: {e}")

#Meter los valores a modificar
input_path=r"H:\unadm\Foto.jpeg" #Ruta de la imagen original
output_path=r"H:\unadm\foto_reducida.jpg" #Ruta de la imagen convertida
compress_image_to_1mb(input_path,output_path)