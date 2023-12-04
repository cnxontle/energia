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

# Diccionario de colores
colores = [simplekml.Color.yellow, simplekml.Color.orange, simplekml.Color.red, simplekml.Color.brown, simplekml.Color.gray, simplekml.Color.pink, simplekml.Color.white, simplekml.Color.orangered, simplekml.Color.green, simplekml.Color.blue, simplekml.Color.purple]

# Opciones del mapa
root = tk.Tk()
app = interfaz.OpcionesMapa(root)
app.iniciar_interfaz()
opciones = app.get_opciones_seleccionadas()

if opciones != []:
    status_filtrar = opciones[0]
    icono_valor = opciones[1]
    tamano_burbuja_valor = opciones[2]
    colorear_por_valor = opciones[3]

    # Cargar el archivo Excel
    script_dir = os.path.dirname(os.path.abspath(__file__))
    destino = os.path.abspath(os.path.join(script_dir, os.pardir))
    RutaMadre = destino + "\\"
    os.chdir(RutaMadre)
    excel_file = RutaMadre + "CUOTA ENERGETICA.xlsm"
    df = pd.read_excel(excel_file)

    # Filtrar por RPU
    for column in df.columns:
        if 'RPU' in column.upper():
            rpu_column = column
            break
    # Verificar si se encontró la columna "RPU"
    if rpu_column is None:
        print("La columna 'RPU' no se encontró en el archivo Excel.")
    else:
        # Filtrar filas donde "RPU" no está vacío
        df = df.dropna(subset=[rpu_column])

    # filtrar por "ESTATUS"
    if status_filtrar == 'Sí':
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
    if tamano_burbuja_valor == '':
        min_tamano = 1
        max_tamano = 2
    else:
        min_tamano = df[tamano_burbuja_valor].min()
        max_tamano = df[tamano_burbuja_valor].max()

    # Crear un objeto KML
    kml = simplekml.Kml()
    
    #Crear diccionario de variables a colorear
    variables_para_colorear = set()
    for index, row in df.iterrows():
        variable_color = row[colorear_por_valor]
        if pd.notna(variable_color):                #si no es nulo
            if variable_color not in variables_para_colorear:
                variables_para_colorear.add(variable_color)
    variables_colorear = list(variables_para_colorear)

    # Iterar a través de las filas del DataFrame
    for index, row in df.iterrows():
        error_vacio = False
        productor = row['NOMBRE']
        try:
            rpu = int(row['RPU'])
        except:
            rpu = row['RPU']
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
        
        if tamano_burbuja_valor == '':
            tamano_punto = 1
        else:
            # Escalar el tamaño de los puntos según la variable tamaño
            variable_tamano = row[tamano_burbuja_valor]
            try:
                tamano_punto = (variable_tamano - min_tamano) / (max_tamano - min_tamano) * (4 - 0.5) + 0.5
            except:
                tamano_punto = 0
            if not pd.notna(variable_tamano):
                error_vacio = True

        if not error_vacio:

            # Colorear los puntos según la variable color
            if colorear_por_valor == '':
                color_punto = 'orangered'
            else:
                variable_color = row[colorear_por_valor]
                if pd.notna(variable_color):
                    color_punto = colores[(variables_colorear.index(variable_color))%11]
                    print(color_punto)

            # Convierte coordenadas DMS a grados decimales usando la función
            try:
                latitud_decimal = dms_to_decimal(latitud_decimal)
                longitud_decimal = dms_to_decimal(longitud_decimal)
            except:
                pass

            if latitud_decimal and longitud_decimal != None:
                # Crear un punto en el KML con coordenadas en grados decimales y aplicar el tamaño
                punto = kml.newpoint(coords=[(longitud_decimal, latitud_decimal)])
                punto.extendeddata.newdata("PRODUCTOR", productor)
                punto.extendeddata.newdata("LOCALIDAD", localidad)
                punto.extendeddata.newdata("RPU", rpu)
                punto.extendeddata.newdata("NO_DE_BOMBAS", str(no_de_bombas))
                punto.extendeddata.newdata("POTENCIA", f'{str(potencia)} HP')
                try:
                    punto.extendeddata.newdata("CUOTA_ENERGETICA",  f'{int(cuota_energetica)} kWh')
                except:
                    punto.extendeddata.newdata("CUOTA_ENERGETICA",  cuota_energetica)
                punto.style.iconstyle.icon.href = icono_valor
                punto.style.iconstyle.scale = tamano_punto
                punto.style.iconstyle.extrude = 0
                punto.style.iconstyle.color = color_punto

    # Guardar el archivo KML
    kml_file = 'puntos.kml'
    kml.save(kml_file)

    print(f'Se ha creado el archivo KML: {kml_file}')