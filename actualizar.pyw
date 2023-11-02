import os
import zipfile
import requests

def descargar_y_descomprimir_zip(url_zip):
    # Descargar el archivo ZIP desde la URL
    response = requests.get(url_zip)
    if response.status_code == 200:
        # Guardar el contenido del ZIP en un archivo local
        with open("programa_actualizado.zip", "wb") as zip_file:
            zip_file.write(response.content)

        # Obtener el directorio del script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        # Calcular el directorio de destino retrocediendo un nivel
        destino = os.path.abspath(os.path.join(script_dir, os.pardir))

        # Descomprimir el archivo ZIP en el directorio de destino
        with zipfile.ZipFile("programa_actualizado.zip", "r") as zip_ref:
            # Extraer el contenido de la carpeta "main" en el directorio de destino
            zip_ref.extractall(destino)

        # Eliminar el archivo ZIP descargado
        os.remove("programa_actualizado.zip")
        print("Programa actualizado correctamente.")
    else:
        print("Error al descargar el archivo ZIP")

if __name__ == "__main__":
    # URL del archivo ZIP en GitHub
    url_zip = "https://github.com/cnxontle/energia/archive/main.zip"

    descargar_y_descomprimir_zip(url_zip)