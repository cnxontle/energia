import pymsgbox
import os
import pandas as pd
from datetime import datetime

def filtrar_filas(columna,filtro):
    global df
    for column in df.columns:
        if columna in column.upper():
            columna_filtro = column
            break
    if columna_filtro is None:
        print(f"La columna '{columna}' no se encontró en el archivo Excel.")
    else:
        if filtro == "N/A":
            df = df.dropna(subset=[columna_filtro])
        else:
            df = df[df[columna_filtro] == filtro]

# Cargar el archivo Excel
script_dir = os.path.dirname(os.path.abspath(__file__))
destino = os.path.abspath(os.path.join(script_dir, os.pardir))
RutaMadre = destino + "\\"
os.chdir(RutaMadre)
excel_file = RutaMadre + "CUOTA ENERGETICA.xlsm"
df = pd.read_excel(excel_file)
filtrar_filas('ESTATUS','N/A')
filtrar_filas('ESTADO DEL PERMISO','VIGENTE')
columnas_objetivo = ['MUNICIPIO', 'LOCALIDAD', 'CLAVE DE REGISTRO (PEUA)', 'NOMBRE', 'CURP', 'RFC', 'TIPO DE PERSONA',
                      'RPU', 'NO. DE CUENTA', 'RMU', 'TIPO DE DOCUMENTO QUE ACREDITA EL USO Y APROVECHAMIENTO DE AGUA',
                      'NUMERO DE FOLIO DEL DOCUMENTO PRESENTADO', 'VIGENCIA DEL TITULO DE CONCESION  (dd/mm/aaaa)',
                      'LATITUD', 'LONGITUD', 'HP DEL EQUIPO REGISTRADO', 'CUOTA ENERGETICA CALCULADA', 'SUPERFICIE DE RIEGO',
                      'CULTIVO','FECHA DE SOLICITUD']

# Crear un nuevo DataFrame con las columnas objetivo
df = df[columnas_objetivo]

# Modificar el formatos de columnas
df['RPU'] = df['RPU'].astype(str)
df['RPU'] = df['RPU'].str.replace('.0', '')
df['NUMERO DE FOLIO DEL DOCUMENTO PRESENTADO'] = df['NUMERO DE FOLIO DEL DOCUMENTO PRESENTADO'].astype(str)
df['FECHA DE SOLICITUD'] = pd.to_datetime(df['FECHA DE SOLICITUD'], errors='coerce')
df['FECHA DE SOLICITUD'] = df['FECHA DE SOLICITUD'].dt.strftime('%d/%m/%Y')
df['MUNICIPIO'] = df['MUNICIPIO'].str.upper()
df['LOCALIDAD'] = df['LOCALIDAD'].str.upper()
df['NOMBRE'] = df['NOMBRE'].str.upper()
df['CURP'] = df['CURP'].str.upper()
df['RFC'] = df['RFC'].str.upper()
df['TIPO DE PERSONA'] = df['TIPO DE PERSONA'].str.upper()
df['NO. DE CUENTA'] = df['NO. DE CUENTA'].str.upper()
df['RMU'] = df['RMU'].str.upper()
df['CULTIVO'] = df['CULTIVO'].str.upper()

# Obtener la fecha actual en formato YYYYMMDD_HHMMSS
fecha_actual = datetime.now().strftime("%d_%m_%Y")
nombre_archivo = f"padron_{fecha_actual}.xlsx"

# Exportar el DataFrame a un archivo Excel
df.to_excel(nombre_archivo, index=False)
pymsgbox.alert('Procedimiento completado!', 'EXPORTAR TABLA')