import os
import zipfile
import requests

def descargar_y_descomprimir_zip(url_zip, destino):
    response = requests.get(url_zip)
    if response.status_code == 200:
        with open("programa_actualizado.zip", "wb") as zip_file:
            zip_file.write(response.content)
        with zipfile.ZipFile("programa_actualizado.zip", "r") as zip_ref:
            zip_ref.extractall(destino)
        os.remove("programa_actualizado.zip")
        print("Programa actualizado correctamente.")
    else:
        print("Error al descargar el archivo ZIP")


#Llamar a la funcion
url_zip = "https://github.com/cnxontle/energia/archive/main.zip"

script_dir = os.path.dirname(os.path.abspath(__file__))
destino = os.path.abspath(os.path.join(script_dir, os.pardir))

descargar_y_descomprimir_zip(url_zip, destino)