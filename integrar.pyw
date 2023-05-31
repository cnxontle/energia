from PyPDF2 import PdfFileMerger, PdfFileReader
import openpyxl
import pymsgbox
import os

# PARAMETROS
RutaExp = __file__.rstrip("integrar.pyw") + "Validacion\\Archivo\\Expedientes\\"
RutaHijo = __file__.rstrip("integrar.pyw") + "Validacion\\Archivo\\"
pdfs_hijo = [archivo for archivo in os.listdir(RutaHijo) if archivo.endswith(".pdf")]
RutaMadre = __file__.rstrip("integrar.pyw")
os.chdir(RutaHijo)
lista_prefijos = []
lista_hijo = []
FilePath = RutaMadre + "CUOTA ENERGETICA.xlsm"
fusionador = PdfFileMerger()
titulos = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "x"]
nombres_titulos = ["Actualizacion", "Biometricos", "PEUA", "Verificacion", "Identificacion", "CURP", "RFC",
				   "Recibo CFE", "Facturas", "Titulo CNA", "Acta Constututiva", "Carta Poder", "Escrituras", "Croquis"]

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

# Lista hijo
for pdf in pdfs_hijo:
	prefijo = pdf[0:4]
	lista_prefijos.append(str(prefijo))
for item in lista_prefijos:
	if item not in lista_hijo:
		lista_hijo.append(item)

# CONSOLIDAR ARCHIVOS
for indice in lista_hijo:
	try:
		fusionador.close()
		fusionador = PdfFileMerger()
		try:
			nombre = rpu[num.index(indice)]  # try
		except:
			nombre = indice
		nombre_archivo_salida = RutaExp + str(nombre) + ".pdf"
		p = 0
		for pdf in sorted(pdfs_hijo):
			if pdf.startswith(indice):
				read_pdf = PdfFileReader(pdf)
				num_pages = read_pdf.getNumPages()
				fusionador.append(pdf)
				titulo = str(pdf[4])
				try:
					titulo_nombre = nombres_titulos[titulos.index(titulo)]  # try
				except:
					titulo_nombre = titulo
				fusionador.addBookmark(titulo_nombre, p)
				p = p + num_pages
		with open(nombre_archivo_salida, 'wb') as salida:
			fusionador.write(salida)
	except:
		pass
fusionador.close()
pymsgbox.alert("Procedimiento Completado!",title='Integrar Expedientes')

