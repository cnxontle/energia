from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import glob
import shutil
import os
import pymsgbox

# PARAMETROS
script_dir = os.path.dirname(os.path.abspath(__file__))
destino = os.path.abspath(os.path.join(script_dir, os.pardir))
RutaRespaldo= destino + "\\Validacion\\Respaldo\\"
RutaMadre= destino + "\\Validacion\\Archivo\\"
RutaTemp= destino + "\\Validacion\\Temp\\"
RutaExp= destino + "\\Validacion\\Archivo\\Expedientes\\"
elementos=[]
documento=("")
prefijo=0
pdfWriter = PdfFileWriter() 
pdfWriter2 = PdfFileWriter()

#FUNCIONES
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
	
def variables(a):
	global elementos
	global documento
	element=a.split(",")
	documento= element[0]
	del element [0]
	for x in element:
		if len(x.split("-"))==1:
			elementos.append(x)
		else:
			y=x.split("-")
			a=int(y[0])
			b=int(y[1])
			c=b-a
			w=[str(a)]
			for i in range(c):
				a=a+1
				
				w.append(str(a))
			for x in w:
				elementos.append(x)

def registro():
	z = open (RutaTemp+'registro.txt','a')
	z.write(prefijo+".pdf,")
	z.close()

#PROCESOS
z =  pymsgbox.confirm(text='¿Que desea Hacer?', title='Edicion de Documentos', buttons=['Restaurar','Borrar','Copiar','Reubicar','Rotar','Fragmentar'])
if z == "Fragmentar":
	a=pymsgbox.prompt('Ingresa expediente, claves y posiciónes, o claves y páginas:',title='Fragmentar Expediente',default='####,ClavePos,Clave,Pag-Pag,Pag,ClavePos...')
	try:
		#DEFINIR LAS VARIABLES
		variables(a)
		elem=[]
		modulo=[]
		elementos2=[]
		recorte(documento)
		error=0
		try:
			pdfbase = open(RutaExp+prefijo+'.pdf','rb')
		except:
			pdfbase = open(RutaExp+documento+'.pdf','rb')
		pdfReader = PdfFileReader(pdfbase)
		pp=pdfReader.numPages
		pp=int(pp)
		t=pp 
		contador=0
		for i in elementos[::-1]:
			if i.isdigit() == True:
				elem.append(i)
				t=int(i)
			elif len(i) ==1:
				elem.append(i)
			else:
				if t==pp:
					u=t-int(i[1::])+1
					for h in range(u):
						elem.append(str(t))
						t=t-1
					elem.append(i[0])
					t=int(i[1::])
				else:
					u=t-int(i[1::])
					for h in range(u):
						t=t-1
						elem.append(str(t))
					elem.append(i[0])
					t=int(i[1::])
		for i in elem [::-1]:
			if i.isdigit()==False:
				contador=contador+1
			if contador <=1:
				modulo.append(i)
			else:
				elementos2.append(modulo)
				modulo=[]
				modulo.append(i)
				contador=1
		elementos2.append(modulo)
		#PROCEDIMIENTO
		for elemento in elementos2:
			try:
				indice= elemento[0]
				del elemento [0]
				paginas=elemento
				pdf5=(RutaTemp+prefijo+indice+'.pdf')
				pdf6=(RutaMadre+prefijo+indice+'.pdf')
				pdfOutputFile = open(pdf5,'wb')
				for i in paginas:
					i=int(i)-1
					pageobj = pdfReader.getPage(i)
					pdfWriter.addPage(pageobj)
				pdfWriter.write(pdfOutputFile)
				pdfOutputFile.close()
				pdfWriter = PdfFileWriter()
				try:
					shutil.copy(pdf6, RutaRespaldo)
				except:
					pass
				shutil.copy(pdf5, RutaMadre)
				os.remove(pdf5)
				z = open (RutaTemp+'registro.txt','a')
				z.write(prefijo+indice+".pdf,")
				z.close()
			except:
				error=1
		pdfbase.close()
		#MENSAJES DE SALIDA
		if error == 0:
			pymsgbox.alert("¡Procedimiento Completado!",title='Fragmentar Expediente')#variable
		elif error == 1:
			pymsgbox.alert("¡El expediente no se fragmentó correctamente, Revisa el número de paginas!",title='Fragmentar Expediente')
	except:
		if a== None:
			pass
		else:
			pymsgbox.alert("¡No se encontro el expediente o hay un error de Sintaxis!",title='Fragmentar Expediente')
	
elif z== 'Reubicar':
	a=pymsgbox.prompt('Ingresa el documento, las paginas, el objetivo y la posición:',title='Reubicar Páginas', default='Documento,Pag-Pag,Pag,DocObjetivo,Pos')
	#DEFINIR VARIABLES
	try:
		variables(a)
		recorte(documento)
		prefijo2=prefijo
		pdf3=(RutaTemp+prefijo+'.pdf')
		pdf4=(RutaTemp+'temp.pdf')
		pdfbase = open(RutaMadre+prefijo+'.pdf','rb')
		pdfReader = PdfFileReader(pdfbase)
		pp=pdfReader.numPages
		pos=elementos[-1]
		if pos.isdigit()==True:
			objetivo=elementos[-2]
			del elementos [-1]
			del elementos [-1]
			paginas=elementos
		else:
			pos=0
			objetivo=elementos[-1]
			del elementos [-1]
			paginas=elementos
		recorte(objetivo) #este es el objetivo (prefijo)
		pdf2=(RutaTemp+prefijo+'.pdf')
		x=1
		if paginas == []:
			for i in range(pp):
				paginas.append(str(x))
				x+=1
	#PROCEDIMIENTO
		pdfOutputFile = open(pdf2,'wb')
		#GENERAR PDF1, seleccionar paginas en objetivo antes de la posicion
		try:
			pdfobjetivo = open(RutaMadre+prefijo+'.pdf','rb')
			pdfReader2 = PdfFileReader(pdfobjetivo)
			pp2=pdfReader2.numPages
			for i in range(pp2):
				pageobj = pdfReader2.getPage(i)
				if i+1 < int(pos):
					pdfWriter.addPage(pageobj)
			pdfWriter.write(pdfOutputFile)
		except:
			pass
		#GENERAR PDF2	seleccionar las paginas a copiar del documento
		for i in range (pp):
			pageobj = pdfReader.getPage(i)
			if str(i+1) in paginas:
				pdfWriter.addPage(pageobj)
		pdfWriter.write(pdfOutputFile)
		pdfbase.close()
		#GENERAR PDF3	seleccionar paginas en objetivo despues de la posicion
		try:
			for i in range(pp2):
				pageobj = pdfReader2.getPage(i)
				if i+1 >= int(pos):
					pdfWriter.addPage(pageobj)
			pdfWriter.write(pdfOutputFile)
		except:
			pass	
		#BORRAR LAS PAGINAS QUE SE COPIARON
		if prefijo2 == prefijo:
			#crear indices para borrar
			p=len(paginas)
			indices=[]
			for i in paginas:
				i= int(i)
				if i < int(pos)+1:
					indices.append(i)
				else:
					i+=p
					indices.append(i)
			#borrar
			pdfOutputFile.close()
			pdfdocumento = open(RutaTemp+prefijo2+'.pdf','rb')
			pdfReader3 = PdfFileReader(pdfdocumento)
			pp3=pdfReader3.numPages
			for i in range(pp3):
				pageobj = pdfReader3.getPage(i)
				if (i+1) not in indices:
					pdfWriter2.addPage(pageobj)
			pdfOutputFile2 = open(pdf4,'wb')
			pdfWriter2.write(pdfOutputFile2)
			pdfOutputFile2.close()
			pdfdocumento.close()
			#eliminar archivo copiado y reemplasar nombre en archivo temporal
			os.remove(RutaTemp+prefijo+'.pdf')
			os.rename(pdf4, pdf3)
		else:
			pdfOutputFile2 = open(pdf3,'wb')
			pdfdocumento = open(RutaMadre+prefijo2+'.pdf','rb')
			pdfReader3 = PdfFileReader(pdfdocumento)
			pp3=pdfReader3.numPages
			for i in range(pp3):
				pageobj = pdfReader3.getPage(i)
				if str(i+1) not in paginas:
					pdfWriter2.addPage(pageobj)
			pdfWriter2.write(pdfOutputFile2)
			pdfOutputFile.close()
			pdfOutputFile2.close()
			pdfdocumento.close()
		pdfpath= (RutaMadre+prefijo+'.pdf')
		pdfpath2= (RutaMadre+prefijo2+'.pdf')
		try:
			shutil.copy(pdfpath, RutaRespaldo)					
		except:
			pass
		shutil.copy(pdf2, RutaMadre)
		os.remove(pdf2)
		registro()
		if prefijo2 == prefijo:
			pass
		else:
			shutil.copy(pdfpath2, RutaRespaldo)   
			pdfbase8 = open(pdf3,'rb')
			pdfReader8 = PdfFileReader(pdfbase8)
			pp8=pdfReader8.numPages
			pdfbase8.close()
			if pp8 == 0:
				os.remove(RutaMadre+prefijo2+'.pdf')
				os.remove(RutaTemp+prefijo2+'.pdf')
			else:
				shutil.copy(pdf3, RutaMadre)
				os.remove(pdf3)
			z = open (RutaTemp+'registro.txt','a')
			z.write(prefijo2+".pdf,")
			z.close()
	#MENSAJES DE SALIDA	
		pymsgbox.alert("Procedimiento Completado!",title='Reubicar Páginas')
	except:
		if a== None:
			pass
		else:
			pymsgbox.alert("¡No se encontro el documento especificado!",title='Reubicar Páginas')	

elif z== 'Copiar':
	a=pymsgbox.prompt('Ingresa el documento, las paginas, el objetivo y la posición:',title='Copiar Páginas', default='Documento,Pag-Pag,Pag,DocObjetivo,Pos')
	#DEFINIR VARIABLES
	try:
		variables(a)
		recorte(documento)
		pdfbase = open(RutaMadre+prefijo+'.pdf','rb')
		pdfReader = PdfFileReader(pdfbase)
		pp=pdfReader.numPages
		pos=elementos[-1]
		if pos.isdigit()==True:
			objetivo=elementos[-2]
			del elementos [-1]
			del elementos [-1]
			paginas=elementos
		else:
			pos=0					
			objetivo=elementos[-1]	
			del elementos [-1]
			paginas=elementos
		recorte(objetivo) #este es el objetivo (prefijo)
		pdf2= (RutaTemp+prefijo+'.pdf')
		x=1
		if paginas == []:
			for i in range(pp):
				paginas.append(str(x))
				x+=1
	#PROCEDIMIENTO
		pdfOutputFile = open(pdf2,'wb')
		#GENERAR PDF1, seleccionar paginas en objetivo antes de la posicion
		try:
			pdfobjetivo = open(RutaMadre+prefijo+'.pdf','rb')
			pdfReader2 = PdfFileReader(pdfobjetivo)
			pp2=pdfReader2.numPages
			for i in range(pp2):
				pageobj = pdfReader2.getPage(i)
				if i+1 < int(pos):
					pdfWriter.addPage(pageobj)
			pdfWriter.write(pdfOutputFile)
		except:
			pass
		#GENERAR PDF2	seleccionar las paginas a copiar del documento
		for i in range (pp):
			pageobj = pdfReader.getPage(i)
			if str(i+1) in paginas:
				pdfWriter.addPage(pageobj)
		pdfWriter.write(pdfOutputFile)
		pdfbase.close()
		#GENERAR PDF3	seleccionar paginas en objetivo despues de la posicion
		try:
			for i in range(pp2):
				pageobj = pdfReader2.getPage(i)
				if i+1 >= int(pos):
					pdfWriter.addPage(pageobj)
			pdfWriter.write(pdfOutputFile)
		except:
			pass	
		pdfOutputFile.close()
		pdfpath= (RutaMadre+prefijo+'.pdf')
		try:
			shutil.copy(pdfpath, RutaRespaldo)#hasta aqui llega bien
		except:
			pass
		shutil.copy(pdf2, RutaMadre)
		os.remove(pdf2)
		registro()
	#MENSAJES DE SALIDA	
		pymsgbox.alert("Procedimiento Completado!",title='Copiar Páginas')
	except:
		if a== None:
			pass
		else:
			pymsgbox.alert("¡No se encontro el documento especificado!",title='Copiar Páginas')	

elif z== "Rotar":
	a=pymsgbox.prompt('Ingresa el documento y las páginas (derecha por default):',title='Rotar Páginas', default='Documento,Pag-Pag,Pag,voltear')
	#DEFINIR VARIABLES
	try:
		variables(a)
		recorte(documento)
		pdf = open(RutaMadre+prefijo+'.pdf','rb')
		pdf2= (RutaTemp+prefijo+'.pdf')
		pdfReader = PdfFileReader(pdf)
		pp=pdfReader.numPages
		paginas=[]
		x=1
		if elementos == []:
			orientacion="derecha"
			for i in range(pp):
				paginas.append(str(x))
				x+=1
		else:	
			orientacion=elementos[-1]
			if orientacion.isdigit()==True:
				orientacion="derecha"
				paginas=elementos
			else:
				del elementos [-1]
				if elementos == []:
					for i in range(pp):
						paginas.append(str(x))
						x+=1
				else:
					paginas=elementos
	#PROCEDIMIENTO
		pdfOutputFile = open(pdf2,'wb')
		for i in range (pp):
			pageobj = pdfReader.getPage(i)
			if str(i+1) not in paginas:
				pass
			else:
				if orientacion == "derecha":
					pageobj.rotateClockwise(90)
				elif orientacion == "voltear":
					pageobj.rotateClockwise(180)
				else:
					pageobj.rotateClockwise(270)
			pdfWriter.addPage(pageobj)
		pdfWriter.write(pdfOutputFile)
		pdf.close()
		pdfOutputFile.close()
		pdfpath= (RutaMadre+prefijo+'.pdf')
		shutil.copy(pdfpath, RutaRespaldo)
		shutil.copy(pdf2, RutaMadre)
		os.remove(pdf2)
		registro()
	#MENSAJES DE SALIDA	
		pymsgbox.alert("Procedimiento Completado!",title='Rotar Páginas')
	except:
		if a== None:
			pass
		else:
			pymsgbox.alert("¡No se encontro el documento especificado!",title='Borrar Páginas')
	
elif z== "Borrar":
	#DEFINIR VARIABLES
	a=pymsgbox.prompt('Ingresa el documento y las páginas que quieras borrar:',title='Borrar Páginas', default='Documento,Pag-Pag,Pag (Todas por default)')
	try:
		variables(a)
		paginas=elementos
		#ELIMINAR EL DOCUMENTO COMPLETO
		if elementos == []:
			recorte(documento)#modificamos la variable prefijo
			pdfpath= (RutaMadre+prefijo+'.pdf')
			shutil.copy(pdfpath, RutaRespaldo)
			os.remove(pdfpath)
			registro()
		#ELIMINAR EL ALGUNAS PAGINAS
		else:
			recorte(documento)#modificamos la variable prefijo
			pdfpath= (RutaMadre+prefijo+'.pdf')
			pdf = open(RutaMadre+prefijo+'.pdf','rb')
			pdfReader = PdfFileReader(pdf)
			pdf2= (RutaTemp+prefijo+'.pdf')
			pp=pdfReader.numPages
			try:
				for i in range(pp):
					if str(i+1) not in paginas:
						pageobj = pdfReader.getPage(i)
						pdfWriter.addPage(pageobj)
						pdfOutputFile = open(pdf2,'wb')
						pdfWriter.write(pdfOutputFile)
				pdf.close()
				pdfOutputFile.close()
				shutil.copy(pdfpath, RutaRespaldo)
				shutil.copy(pdf2, RutaMadre)
				os.remove(pdf2)
				registro()
			except:
				pdf.close()
				shutil.copy(pdfpath, RutaRespaldo)
				os.remove(pdfpath)
				registro()
		#MENSAJES DE SALIDA
		pymsgbox.alert("Procedimiento Completado!",title='Borrar Páginas')
	except:
		if a== None:
			pass
		else:
			pymsgbox.alert("¡No se encontro el documento especificado!",title='Borrar Páginas')
elif z== "Restaurar":
	a=pymsgbox.prompt('Ingresa los documentos que desees restaurar:',title='Restaurar Documentos', default='Documento,Documento,Documento...')
	docs=a.split(",")
	try:	
		for i in docs:
			i=recorte(i)
			try:
				shutil.copy(RutaMadre+prefijo+".pdf", RutaTemp)
			except:
				pass
			shutil.copy(RutaRespaldo+prefijo+".pdf", RutaMadre)
			try:
				shutil.copy(RutaTemp+prefijo+".pdf", RutaRespaldo)
			except:
				os.remove(RutaRespaldo+prefijo+".pdf")
			registro()
		for pdf in glob.glob(RutaTemp+'*.pdf'):
			os.remove(pdf)
		#MENSAJES DE SALIDA	
		pymsgbox.alert("Procedimiento Completado!",title='Restaurar Páginas')
	except:
		pymsgbox.alert("¡No se encontro el documento especificado!",title='Restaurar Páginas')			
		for pdf in glob.glob(RutaTemp+'*.pdf'):
			os.remove(pdf)
