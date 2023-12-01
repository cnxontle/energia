import simplekml
import pandas as pd
import re
import os
import interfaz
import tkinter as tk

def dms_to_decimal(dms):
    # Manejar valores vacíos o nulos
    if pd.isna(dms) or dms == '':
        return None
    
    # Elimina espacios en blanco y comillas dobles
    dms = dms.replace(' ', '').replace('"', '')

    # Divide la cadena en partes en función de los diferentes delimitadores
    parts = re.split(r'[°\'"]+', dms)

    # Convierte grados y minutos a números
    degrees = float(parts[0])
    minutes = float(parts[1])
    seconds = float(parts[2])

    if degrees > 70:
        decimal = -(degrees + minutes / 60 + seconds / 3600)
    elif degrees < 0:
        decimal = degrees - minutes / 60 - seconds / 3600
    else:
        decimal = degrees + minutes / 60 + seconds / 3600
    return decimal

# Opciones del mapa
root = tk.Tk()
app = interfaz.OpcionesMapa(root)
app.iniciar_interfaz()
opciones = app.get_opciones_seleccionadas()

if opciones != []:
    status_filtrar = opciones[0]
    tamano_burbuja_valor = opciones[1]
    colorear_por_valor = opciones[2]

    # Cargar el archivo Excel
    script_dir = os.path.dirname(os.path.abspath(__file__))
    destino = os.path.abspath(os.path.join(script_dir, os.pardir))
    RutaMadre = destino + "\\"
    os.chdir(RutaMadre)
    excel_file = RutaMadre + "CUOTA ENERGETICA.xlsm"


    df = pd.read_excel(excel_file)

    # Buscar la columna "ESTATUS"
    estatus_column = None
    for column in df.columns:
        if 'ESTATUS' in column.upper():
            estatus_column = column
            break

    # Verificar si se encontró la columna "ESTATUS"
    if estatus_column is None:
        print("La columna 'ESTATUS' no se encontró en el archivo Excel.")
    else:
        # Filtrar filas donde "ESTATUS" no está vacío
        df = df.dropna(subset=[estatus_column])

        # Calcular los valores mínimos y máximos de "CUOTA_ENERGETICA"
        min_cuota = df['CUOTA ENERGETICA'].min()
        max_cuota = df['CUOTA ENERGETICA'].max()

        # Crear un objeto KML
        kml = simplekml.Kml()

        # Iterar a través de las filas del DataFrame
        for index, row in df.iterrows():
            productor = row['NOMBRE']
            rpu = int(row['RPU'])
            no_de_bombas = row['NO. DE BOMBAS']
            try:
                no_de_bombas = int(no_de_bombas)
            except:
                pass
            potencia = row['HP DEL EQUIPO REGISTRADO']
            cuota_energetica = row['CUOTA ENERGETICA']
            localidad = row['LOCALIDAD']
            latitud_decimal = row['LATITUD']
            longitud_decimal = row['LONGITUD']
            
            # Escalar el tamaño de los puntos según la cuota energética
            tamaño_punto = (cuota_energetica - min_cuota) / (max_cuota - min_cuota) * (4 - 0.5) + 0.5

            # Convierte coordenadas DMS a grados decimales usando la función
            latitud_decimal = dms_to_decimal(latitud_decimal)
            longitud_decimal = dms_to_decimal(longitud_decimal)

            # Crear un punto en el KML con coordenadas en grados decimales y aplicar el tamaño
            punto = kml.newpoint(coords=[(longitud_decimal, latitud_decimal)])
            punto.extendeddata.newdata("PRODUCTOR", productor)
            punto.extendeddata.newdata("LOCALIDAD", localidad)
            punto.extendeddata.newdata("RPU", rpu)
            punto.extendeddata.newdata("NO_DE_BOMBAS", str(no_de_bombas))
            punto.extendeddata.newdata("POTENCIA", f'{str(potencia)} HP')
            punto.extendeddata.newdata("CUOTA_ENERGETICA",  f'{int(cuota_energetica)} kWh')
            punto.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
            punto.style.iconstyle.scale = tamaño_punto
            punto.style.iconstyle.extrude = 0
            punto.style.iconstyle.color = simplekml.Color.red

        # Guardar el archivo KML
        kml_file = 'puntos.kml'
        kml.save(kml_file)

        print(f'Se ha creado el archivo KML: {kml_file}')
