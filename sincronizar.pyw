import glob
import shutil
import os
import pymsgbox

# PARAMETROS
RutaMadre=__file__.rstrip("sincronizar.pyw")+"Validacion\\Archivo\\"
RutaTemp= __file__.rstrip("sincronizar.pyw")+"Validacion\\Temp\\"
RutaMadre2="Validacion\\Archivo\\"
RutaTemp2="Validacion\\Temp\\"
excel=__file__.rstrip("sincronizar.pyw")+"CUOTA ENERGETICA.xlsm"
excel2="CUOTA ENERGETICA.xlsm"
complemento="Validacion\\Archivo\\"
registros=[]
prefijo=[]
registros2=[]
registros3=[]
registros4=[]
vecinos=[]
master="no"

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
 
#LEER NODOS
f = open (RutaTemp+'nodos.txt','r')
nodos=[]
for line in f:
	line=line.strip()
	line=line.split(",")
	nodos.append(line)
f.close()

#SOY ADMINISTRADOR?              
for nodo in nodos:
	direccion=nodo[1]
	n=nodo[0]
	if direccion == "self":
		z=1
		identidad=n
	if n == "nodo1":
		zz=1
	if z==zz:
		master="si"

#VECINOS
for nodo in nodos:
	direccion=nodo[1]
	n=nodo[0]
	if direccion == "self":
		pass
	else:
		vecinos.append(n)
		
#BUSCAR Y RECIBIR ACTUALIZACIONES...
for nodo in nodos:
	direccion=nodo[1]
	n=nodo[0] 
	if direccion == "self":
		pass
	else:
 
#RECIBIR DESDE EL REGISTROS PENDIENTES
		try:
			f = open (direccion+RutaTemp2+identidad+'.txt','r')
			g=f.readlines()
			h=str(g)[2:-3]
			j=h.split(",")
			for k in j:
				recorte(k)
				prefijo=prefijo.lower()+".pdf"
				registros3.append(prefijo)
			documentos3=set(registros3)
			f.close()
			for i in documentos3:
				try:
					shutil.copy(direccion+RutaMadre2+i, RutaMadre)
				except:
					try:			
						os.remove(RutaMadre+i)
					except:
						pass	
			os.remove(direccion+RutaTemp2+identidad+'.txt')    #no escribo registro por que el archivo pendiente solo es para mi, y ya no lo pienso compartir
		except:
			pass

#RECIBIR DESDE LOS REGISTROS LOCALES 
for nodo2 in nodos:
	direccion=nodo2[1]
	n2=nodo2[0] 
	if direccion == "self":
		pass
	else: 
		try:
			ff = open (direccion+RutaTemp2+'registro.txt','r')
			g=ff.readlines()
			h=str(g)[2:-3]
			j=h.split(",")
			for k in j:
				recorte(k)
				prefijo=prefijo.lower()+".pdf"
				registros4.append(prefijo)
			documentos4=set(registros4)
			ff.close()
			for i in documentos4:
				try:
					shutil.copy(direccion+RutaMadre2+i, RutaMadre)
				except:
					try:			
						os.remove(RutaMadre+i)
					except:
						pass	
				if i == ".pdf":
					pass
				else:	
					for nod in vecinos:						#guardo un registro personalizado para mis nodos vecinos, excepto el nodo de origen
						if nod==n2:
							pass
						else:
							q = open(RutaTemp+nod+'.txt','a')
							q.write(i+",")
							q.close()
				i = ".pdf"
			ww=open(direccion+RutaTemp2+'registro.txt','w')
			ww.close()
		except:
			pass

#LEER REGISTRO LOCAL 
l = open (RutaTemp+'registro.txt','r')
m=l.readlines()
n=str(m)[2:-3]
o=n.split(",") 
for k in o:
	recorte(k)
	prefijo=prefijo.lower()+".pdf"
	registros.append(prefijo)
documentos=set(registros) 
l.close()

#EMITIR CAMBIOS
for nodo in nodos:
	direccion=nodo[1]
	n=nodo[0]
	if direccion == "self":	
		pass
	else:
		Ruta1=direccion+complemento
		Ruta2=direccion
		#EMITIR DESDE EL REGISTRO PENDIENTE
		try:
			f = open (RutaTemp+n+'.txt','r')
			g=f.readlines()
			h=str(g)[2:-3]
			j=h.split(",")
			for k in j:
				recorte(k)
				prefijo=prefijo.lower()+".pdf"
				registros2.append(prefijo)
			documentos2=set(registros2)
			for i in documentos2:
				try:
					shutil.copy(RutaMadre+i, Ruta1)
				except:
					try:			
						os.remove(Ruta1+i)
					except:
						pass	
			f.close()
			os.remove(RutaTemp+n+'.txt')
		except:
			pass
		#EMITIR DESDE REGISTRO LOCAL
		try:
			if master=="si":
				shutil.copy(excel, Ruta2)		#si soy administrador voy a copiar mi excel en todas partes
			for i in documentos:
				shutil.copy(RutaMadre+i, Ruta1)
				try:
					shutil.copy(RutaMadre+i, Ruta1)
				except:
					try:			
						os.remove(Ruta1+i)
					except:
						pass	
		except:
			cadena=",".join(documentos)
			if cadena == ".pdf":
				pass
			else:
				pymsgbox.alert("no se pudo sincronizar en "+n,title='Sincronizar Documentos')
				f = open (RutaTemp+n+'.txt','a')
				f.write(cadena+",")
				f.close()
				#try:
				#	for vecino in vecinos:
				#		shutil.copy(RutaTemp+n+'.txt', vecino+RutaTemp2)
				#except:
				#	pass
				
m = open (RutaTemp+'registro.txt','w')
m.close()
pymsgbox.alert("Procedimiento completado",title='Sincronizar Documentos')
