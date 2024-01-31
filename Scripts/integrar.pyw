from PyPDF2 import PdfFileMerger, PdfFileReader
import openpyxl
import pymsgbox
import os
import tkinter as tk
from tkinter import ttk

# PARAMETROS
script_dir = os.path.dirname(os.path.abspath(__file__))
destino = os.path.abspath(os.path.join(script_dir, os.pardir))
RutaExp = destino + "\\Validacion\\Archivo\\Expedientes\\"
RutaHijo = destino + "\\Validacion\\Archivo\\"
pdfs_hijo = [archivo for archivo in os.listdir(RutaHijo) if archivo.endswith(".pdf")]
RutaMadre = destino + "\\"
os.chdir(RutaHijo)
p = 0
lista_prefijos = []
FilePath = RutaMadre + "CUOTA ENERGETICA.xlsm"
fusionador = PdfFileMerger()
titulos = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "x"]
nombres_titulos = ["Actualizacion", "Biometricos", "PEUA", "Verificacion", "Identificacion", "CURP", "RFC",
				   "Recibo CFE", "Facturas", "Titulo CNA", "Acta Constututiva", "Carta Poder", "Escrituras", "Croquis"]

# Borramos los archivos de la carpeta de expedientes
for archivo in os.listdir(RutaExp):
	os.remove(RutaExp + archivo)

#FUNCION PARA EXTRAER titulos Y REEMPAZAR POR nombre_titulos
def agregar_bookmark(fusionador, pdf, p, num_pages, titulos, nombres_titulos):
    titulo = str(pdf[4])
    try:
        titulo_nombre = nombres_titulos[titulos.index(titulo)]
    except:
        titulo_nombre = titulo
    fusionador.addBookmark(titulo_nombre, p)
    p = p + num_pages
    return p

# LEER DATOS DE EXCEL
try:
	df = openpyxl.load_workbook(FilePath, data_only=True)
	hoja = df['DATOS']
	celdas = hoja['B1':'B10000']
	celdas2 = hoja['A1':'A10000']
	num = []
	rpu = []
	#esta fila se va a agregar a la lista num
	for fila in celdas:
		if fila[0].value != None:
			x = str(fila[0].value)
			if len(x) < 4:
				x = '0' * (4 - len(x)) + x
			num.append(x)
		else:
			break
	#esta fila se va a agregar a la lista rpu
	for fila in celdas2:
		if fila[0].value != None:
			rpu.append(fila[0].value)
		else:
			break
except:
	pass

# Lista prefijos
for pdf in pdfs_hijo:
	prefijo = pdf[0:4]
	lista_prefijos.append(str(prefijo))

# CONFIGURAR BARRA DE PROGRESO
ventana = tk.Tk()
ventana.title("Integrando expedientes...")
ventana.geometry("310x110")
barra_progreso = ttk.Progressbar(ventana, length=280)
barra_progreso.pack(pady=20)
barra_progreso.configure(mode="determinate", maximum=len(lista_prefijos))
etiqueta_mensaje = tk.Label(ventana, text="Para cancelar el proceso, cierra la ventana...      \n\n", font=("Arial", 10), anchor="w")
etiqueta_mensaje.pack()

# CONSOLIDAR ARCHIVOS
for pdf in sorted(pdfs_hijo):
	try:
		prefijo_adelantado=lista_prefijos[1]
	except:
		prefijo_adelantado="$$$$"
	lista_prefijos.pop(0)
	read_pdf = PdfFileReader(pdf)
	num_pages = read_pdf.getNumPages()
	fusionador.append(pdf)
	p = agregar_bookmark(fusionador, pdf, p, num_pages, titulos, nombres_titulos)
	
	#renombramos y escribimos el archivo
	if not pdf.startswith(prefijo_adelantado):
		try:
			nombre = rpu[num.index(pdf[:4])]
		except:
			nombre = pdf[:4]
		nombre_archivo_salida = RutaExp + str(nombre) + ".pdf"
		p = 0
		with open(nombre_archivo_salida, 'wb') as salida:
			fusionador.write(salida)
		#reiniciamos el fusionador
		fusionador.close()
		fusionador = PdfFileMerger()
	try:
		barra_progreso['value'] += 1
	except:
		break
	ventana.update()

# Cerrar la ventana después de que se complete el proceso
fusionador.close()
try:
	ventana.destroy()
	pymsgbox.alert("Procedimiento Completado!",title='Integrar Expedientes')
except:
	pass
ventana.mainloop()