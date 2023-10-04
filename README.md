# energia
Integración de documentos en pdf

Se rquieren las siguientes librerias

pip install PyPDF2==1.26.0

pip install pymsgbox

pip install path.py

pip install wheel

pip install openpyxl

pip install simplekml

pip install pandas

nos vamos al editor de registro, para esto buscamos ejecutamos en windows la palabra regedit
al abrir regedit vamos abrimos el arbol HKEY_CLASSES_ROOT y buscamos la clave que dice Python file, seleccionamos la clave y agregamos un nuevo valor DWORD (32 bits) lo renombramos con el nombre EditFlags, con el boton derecho del raton seleccionamos la opcion modificar e ingresamos un valor de 10000 hexadeciimal.
este mismo procedimiento lo hacemos en la clave Python.NoConFile


Con esto todo debe funcionar bien, en caso de que los botones del documento de excel sigan sin funcionar vamos a ejecutar el archivo MSO_Hyperlinks.CMD que se encuentra en esta carpeta
