import simplekml
import pandas as pd
import re
import os
import interfaz
import tkinter as tk
import pymsgbox

def filtrar_columnas(columna):
    global df
    for column in df.columns:
        if columna in column.upper():
            columna_filtro = column
            break
    # Verificar si se encontró la columna
    if columna_filtro is None:
        print(f"La columna '{columna}' no se encontró en el archivo Excel.")
    else:
        # Filtrar filas donde la columna no está vacía
        df = df.dropna(subset=[columna_filtro])

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

# Contadores de errores
errores_coordenadas = 0
errores_vacios = 0

# Diccionario de colores
colores = ['e60099ff','e66666ff','e600ccff','e63333ff','ff99ccff','ff9999ff','e66699ff','ff4c76a6', 'e63366ff','ff6699ff']
leyenda = ['NOMBRE', 'RPU', 'SUPERFICIE DE RIEGO', 'NO. DE BOMBAS', 'HP DEL EQUIPO REGISTRADO', 'CUOTA ENERGETICA', 'LOCALIDAD']

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

    # Medida de la leyenda adicional
    if tamano_burbuja_valor == 'SUPERFICIE DE RIEGO':
        medida = 'ha'
    elif tamano_burbuja_valor == 'CUOTA ENERGETICA':
        medida = 'kWh'
    elif tamano_burbuja_valor == 'CONSUMO ANUAL KWH':
        medida = 'kWh'
    elif tamano_burbuja_valor == 'kWh POR HECTAREA':
        medida = 'kWh/ha'
    elif tamano_burbuja_valor == 'VOLUMEN CONCECIONADO':
        medida = 'm3'
    elif tamano_burbuja_valor == 'GASTO ANUAL M3':
        medida = 'm3'
    elif tamano_burbuja_valor == 'CONSUMO ENERGETICO ENTRE VOLUMEN':
        medida = 'kWh/m3'
    elif tamano_burbuja_valor == 'SUPRIEGO ENTRE SUPBEN':
        medida = ''
    elif tamano_burbuja_valor == 'APROVECHAMIENTO DE LA CUOTA':
        medida = ''
    else:
        medida = ''

    # Cargar el archivo Excel
    script_dir = os.path.dirname(os.path.abspath(__file__))
    destino = os.path.abspath(os.path.join(script_dir, os.pardir))
    RutaMadre = destino + "\\"
    os.chdir(RutaMadre)
    excel_file = RutaMadre + "CUOTA ENERGETICA.xlsm"
    df = pd.read_excel(excel_file)

    df["CONSUMO ENERGETICO ENTRE VOLUMEN"] = df["CONSUMO ANUAL KWH"] / df["VOLUMEN CONCECIONADO"].replace(0, pd.NA)
    df["kWh POR HECTAREA"] = df["CONSUMO ANUAL KWH"] / df["SUPERFICIE DE RIEGO"].replace(0, pd.NA)
    df["APROVECHAMIENTO DE LA CUOTA"] = df["CONSUMO ANUAL KWH"] / df["CUOTA ENERGETICA CALCULADA"].replace(0, pd.NA)
    df["SUPRIEGO ENTRE SUPBEN"] = df["SUPERFICIE DE RIEGO"] / df["SUP BENEFICIADA"].replace(0, pd.NA)

    # Filtrar Columnas
    filtrar_columnas('RPU')
    if status_filtrar == 'Sí':
        filtrar_columnas('ESTATUS')
    filtrar_columnas('LATITUD')
    filtrar_columnas('LONGITUD')
    filtrar_columnas('CARPETA')

    # Calcular los valores mínimos y máximos de "variable tamaño"
    if tamano_burbuja_valor == '':
        min_tamano = 1
        max_tamano = 2
    else:
        try:
            min_tamano = df[tamano_burbuja_valor].min()
            max_tamano = df[tamano_burbuja_valor].max()
        except:
            pymsgbox.alert(f"Se encontró un error en los valores de la columna: '{tamano_burbuja_valor}'", "Error")
            exit()

    # Crear un objeto KML
    kml = simplekml.Kml()
    
    #Crear diccionario de variables a colorear
    if colorear_por_valor != '':
        variables_para_colorear = set()
        for index, row in df.iterrows():
            variable_color = row[colorear_por_valor]
            if pd.notna(variable_color):                
                if variable_color not in variables_para_colorear:
                    variables_para_colorear.add(variable_color)
        variables = list(variables_para_colorear)
        variables_colorear = sorted(variables)

    # Iterar a través de las filas del DataFrame
    for index, row in df.iterrows():
        error_vacio = False
        productor = row['NOMBRE']
        
        try:
            rpu = int(row['RPU'])
        except:
            rpu = row['RPU']
        superficie_de_riego = row['SUPERFICIE DE RIEGO']
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
        
        # Leyenda adicional 1
        if tamano_burbuja_valor not in leyenda and tamano_burbuja_valor != '':
            leyenda1 = row[tamano_burbuja_valor]
        
        # Leyenda adicional 2
        if colorear_por_valor not in leyenda and colorear_por_valor != '':
            if opciones[5].startswith('Checklist') and opciones[5] != 'Checklist_Verificacion':
                leyenda_p = row[colorear_por_valor]
                if pd.notna(leyenda_p):
                    leyenda2 = row[colorear_por_valor]
                else:
                    leyenda2 = 'CORRECTO'
            else:
                leyenda2 = row[colorear_por_valor]

        # Calcular el tamaño del punto
        if tamano_burbuja_valor == '':
            tamano_punto = 1
        else:
            variable_tamano = row[tamano_burbuja_valor]
            try:
                tamano_punto = (variable_tamano - min_tamano) / (max_tamano - min_tamano) * (4 - 0.5) + 0.5
            except:
                tamano_punto = 0
            if not pd.notna(variable_tamano):
                error_vacio = True
                errores_vacios += 1

        # Colorear los puntos según la variable color
        if not error_vacio:
            if colorear_por_valor == '':
                color_punto = 'e63333ff'
            else:
                variable_color = row[colorear_por_valor]
                if pd.notna(variable_color):
                    color_punto = colores[(variables_colorear.index(variable_color))%10]
                else:
                    color_punto = simplekml.Color.gray

            # Convierte coordenadas DMS a grados decimales usando la función
            try:
                latitud_decimal = dms_to_decimal(latitud_decimal)
                longitud_decimal = dms_to_decimal(longitud_decimal)
            except:
                errores_coordenadas += 1
                pass
                
            if latitud_decimal != None and longitud_decimal != None:
                # Crear un punto en el KML con coordenadas en grados decimales y aplicar el tamaño
                punto = kml.newpoint(coords=[(longitud_decimal, latitud_decimal)])
                punto.extendeddata.newdata("PRODUCTOR", productor)
                punto.extendeddata.newdata("LOCALIDAD", localidad)
                punto.extendeddata.newdata("RPU", rpu)
                try:
                    punto.extendeddata.newdata("SUPERFICIE_DE_RIEGO", f'{int(superficie_de_riego)} ha')
                except:
                    punto.extendeddata.newdata("SUPERFICIE_DE_RIEGO", superficie_de_riego)
                punto.extendeddata.newdata("NO_DE_BOMBAS", str(no_de_bombas))
                punto.extendeddata.newdata("POTENCIA", f'{str(potencia)} HP')
                try:
                    punto.extendeddata.newdata("CUOTA_ENERGETICA",  f'{int(cuota_energetica)} kWh')
                except:
                    punto.extendeddata.newdata("CUOTA_ENERGETICA",  cuota_energetica)
                if tamano_burbuja_valor not in leyenda and tamano_burbuja_valor != '':
                    if tamano_burbuja_valor == "APROVECHAMIENTO DE LA CUOTA" or tamano_burbuja_valor == "CONSUMO ENERGETICO ENTRE VOLUMEN" or tamano_burbuja_valor == "SUPRIEGO ENTRE SUPBEN":
                        try: 
                            if tamano_burbuja_valor == "APROVECHAMIENTO DE LA CUOTA":
                                resultado_formateado = "{:.2f}".format(float(leyenda1) * 100)
                                punto.extendeddata.newdata(opciones[4].upper(),  f'{resultado_formateado} %')
                            elif tamano_burbuja_valor == "CONSUMO ENERGETICO ENTRE VOLUMEN":
                                resultado_formateado = "{:.2f}".format(float(leyenda1))
                                punto.extendeddata.newdata(opciones[4].upper(),  f'{resultado_formateado} kWh/m3')
                            elif tamano_burbuja_valor == "SUPRIEGO ENTRE SUPBEN":
                                resultado_formateado = "{:.2f}".format(float(leyenda1))
                                punto.extendeddata.newdata(opciones[4].upper(),  f'{resultado_formateado}')
                        except: 
                            punto.extendeddata.newdata(opciones[4].upper(), leyenda1)
                    else:
                        try:
                            punto.extendeddata.newdata(opciones[4].upper(),  f'{int(leyenda1)} {medida}')
                        except:
                            punto.extendeddata.newdata(opciones[4].upper(), leyenda1)
                if colorear_por_valor not in leyenda and colorear_por_valor != '':
                    punto.extendeddata.newdata(opciones[5].upper(), leyenda2)
                punto.style.iconstyle.icon.href = icono_valor
                punto.style.iconstyle.scale = tamano_punto
                punto.style.iconstyle.extrude = 0
                punto.style.iconstyle.color = color_punto

    # Guardar el archivo KML
    nombre1 = opciones[4].lower()
    nombre2 = opciones[5].lower()
    kml_file = f'mapa_{nombre1}_{nombre2}.kml'
    kml_path = os.path.join("Mapas", kml_file)
    kml.save(kml_path)

    # Mostrar mensaje de finalización
    mensaje = "Procedimiento Completado!\n\n"

    if errores_vacios == 0 and errores_coordenadas == 0:
        mensaje
    elif errores_vacios == 0 and errores_coordenadas != 0:
        mensaje += f"Se encontraron {errores_coordenadas} coordenadas inválidas."
    elif errores_vacios != 0 and errores_coordenadas == 0:
        mensaje += f"Se encontraron {errores_vacios} valores vacíos."
    else:
        mensaje += f"Se encontraron {errores_coordenadas} coordenadas inválidas y {errores_vacios} valores vacíos."

    pymsgbox.alert(mensaje, title='Generacion de Mapa')