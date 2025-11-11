from PyPDF2 import PdfFileMerger, PdfFileReader
import glob
import shutil
import os
import pymsgbox

# PARAMETROS
script_dir = os.path.dirname(os.path.abspath(__file__))
destino = os.path.abspath(os.path.join(script_dir, os.pardir))
RutaTemp= destino +"\\Validacion\\Temp\\"
RutaRespaldo= destino +"\\Validacion\\Respaldo\\"
RutaErrores= destino +"\\Validacion\\Errores\\"
RutaMadre= destino +"\\Validacion\\Archivo\\"
RutaHijo= destino +"\\Validacion\\"
os.chdir(RutaHijo)
pdfs_hijo = [archivo for archivo in os.listdir(RutaHijo) if archivo.endswith(".pdf")]
lista_prefijos = []
lista_hijo= []
lista_madre= []
registro=[]
prefijo=0
error=0
fusionador = PdfFileMerger()

#Funcion Recorte
def recorte (y):
    global prefijo
    y=str(y)
    z=0
    for i in y:
        if i.isdigit() == True:
            z=z+1
        else:
            break
    if z == 1:
        prefijo="000"+y[0:2]
    elif z == 2:
        prefijo="00"+y[0:3]
    elif z == 3:
        prefijo="0"+y[0:4]
    else:
        prefijo=y[0:5]

#Lista hijo
for pdf in pdfs_hijo:
    recorte(pdf)   
    lista_prefijos.append(str(prefijo))
for item in lista_prefijos:
    if item not in lista_hijo:
        lista_hijo.append(item)

#Lista Madre
for pdf in os.listdir(RutaMadre):
    if pdf.endswith(".pdf"):
        recorte(pdf)   
        lista_madre.append(str(prefijo))
        
#CONSOLIDAR DUPLICADOS
for indice in lista_hijo:
    fusionador.close()
    fusionador = PdfFileMerger()
    nombre_archivo_salida = RutaTemp+indice+".pdf"
    for pdf in sorted(pdfs_hijo):
        recorte(pdf) 
        if prefijo.startswith(indice):
            try:
                fusionador.append(pdf)
            except:
                #error: startxref not found
                error=error+1
                shutil.copy(pdf, RutaErrores) 
                pass
    with open(nombre_archivo_salida, 'wb') as salida:
        fusionador.write(salida)
fusionador.close()
fusionador = PdfFileMerger()

#REEMPLAZAR CONSOLIDADOS DESDE RUTA TEMP
for py_file in glob.glob(RutaHijo+'*.pdf'):
    os.remove(py_file)
contenidos=[archivo for archivo in os.listdir(RutaTemp) if archivo.endswith(".pdf")]
for elemento in contenidos:
    shutil.copy(RutaTemp+elemento, RutaHijo)  
for pdf in glob.glob(RutaTemp+'*.pdf'):
    os.remove(pdf)

#GUARDAR RESPALDO
for indice in lista_hijo:
    contenidos=[archivo for archivo in os.listdir(RutaMadre) if archivo.startswith(indice)]
    for elemento in contenidos:
        shutil.copy(RutaMadre+elemento, RutaRespaldo)  
    
#LISTAR DE ARCHIVOS PARA LA ACTUALIZACION    
pdfs_hijo = [archivo for archivo in os.listdir(RutaHijo) if archivo.endswith(".pdf")]
for pdf in pdfs_hijo:
    fusionador.close()
    fusionador = PdfFileMerger()
    recorte(pdf)  
    nombre_archivo_salida = RutaTemp+prefijo+".pdf"
        
    #ACTUALIZAR DESDE RUTA MADRE Y ESCRIBIR EN CARPETA TEMP
    for archivo in os.listdir(RutaMadre):
        if archivo.endswith(".pdf")and archivo.startswith(prefijo):
            fusionador.append(pdf)
            fusionador.append(PdfFileReader(RutaMadre+archivo))
            with open(nombre_archivo_salida, 'wb') as salida:
                fusionador.write(salida)
        
    #INCLUIR DOCUMENTOS NUEVOS
    if prefijo not in lista_madre:
        fusionador.append(pdf)
        with open(nombre_archivo_salida, 'wb') as salida:
            fusionador.write(salida)
fusionador.close()
fusionador = PdfFileMerger()
                    
#SOBREESCRIBIR EN RUTA MADRE
contenidos=[archivo for archivo in os.listdir(RutaTemp) if archivo.endswith(".pdf")]
for elemento in contenidos:
    shutil.copy(RutaTemp+elemento, RutaMadre)
    registro.append(elemento)

#GUARDAR REGISTRO DE EVENTOS
os.chdir(RutaTemp)
f = open ('registro.txt','a')
if str(registro) != "[]":
	cadena=",".join(registro)+","
	f.write(cadena)                  
f.close()

#BORRAR RESIDUOS
for py_file in glob.glob(RutaHijo+'*.pdf'):
    os.remove(py_file)
for pdf in glob.glob(RutaTemp+'*.pdf'):
    os.remove(pdf)

#MENSAJES DE SALIDA
if error == 0:
    pymsgbox.alert('Los documentos se actualizarón correctamente!','CARPETA ARCHIVO')
else:
    pymsgbox.alert('Algunos documentos no se pudieron procesar!','CARPETA ERRORES')    
