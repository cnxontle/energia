# Guia de Configucación

Este proyecto requiere la instalación de varias bibliotecas de Python y algunas configuraciones específicas para habilitar los macros en Excel y eliminar los mensajes de advertencia en los vínculos. Sigue estos pasos para configurar el proyecto correctamente.


## Instalación de Bibliotecas

Asegúrate de tener las siguientes bibliotecas instaladas en su entorno de Python. Puedes instalarlas utilizando pip:

```bash
pip install PyPDF2==1.26.0

pip install pymsgbox

pip install path.py

pip install wheel

pip install openpyxl

pip install simplekml

pip install pandas

pip install xlrd --upgrade
```


## Habilitar Macros


Para habilitar los macros en Excel, agrega la carpeta del proyecto a las "Ubicaciones de Confianza" dentro de las opciones de Excel.


## Eliminar Mensajes de Advertencia en Vínculos

Para eliminar los mensajes de advertencia en los vínculos, sigue estos pasos:

a. Abre el Editor de Registro. Para hacerlo, ejecuta "regedit" en Windows.

b. Navega al árbol de registros HKEY_CLASSES_ROOT.

c. Busca la clave que dice "Python.NoConFile".

d. Selecciona la clave y agrega un nuevo valor DWORD (32 bits) con el nombre "EditFlags".

e. Haz clic derecho en el valor recién creado y selecciona la opción "Modificar". Ingresa un valor de 10000 en hexadecimal.



## Verificar los Botones en el Documento de Excel

Si los botones en el documento de Excel aún no funcionan después de realizar las configuraciones anteriores intenta lo siguiente:

a. ejecuta el archivo "MSO_Hyperlinks.CMD" que se encuentra en la carpeta del proyecto.

b. Navega hasta la clave: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Security.

c. Si la clave "Security" no existe bajo "Common", créala.

d. Busca o agrega una clave de tipo DWORD con el nombre "DisableHyperlinkWarning".

e. Haz doble clic en la clave "DisableHyperlinkWarning" y establece su valor en 1 (uno).

Con estos pasos, ya deberías tener configurado correctamente el proyecto "Energía" y poder trabajar sin problemas. 





