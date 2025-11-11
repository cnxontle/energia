import tkinter as tk
from tkinter import ttk

class OpcionesMapa:
    def __init__(self, master):
        self.master = master
        master.title("Opciones del Mapa")
        master.geometry("327x172+50+50")

        # Lista para almacenar las opciones seleccionadas
        self.opciones_seleccionadas = []
        self.directorio_iconos = ['http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png','http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png','http://maps.google.com/mapfiles/kml/shapes/open-diamond.png','http://maps.google.com/mapfiles/kml/shapes/donut.png','http://maps.google.com/mapfiles/kml/shapes/polygon.png','http://maps.google.com/mapfiles/kml/shapes/triangle.png']
        self.columnas_tamano = ["", "SUPERFICIE DE RIEGO", "CUOTA ENERGETICA CALCULADA", "APROVECHAMIENTO DE LA CUOTA", "CONSUMO ANUAL KWH", "kWh POR HECTAREA", "VOLUMEN CONCECIONADO", "GASTO ANUAL M3", "CONSUMO ENERGETICO ENTRE VOLUMEN", "SUPRIEGO ENTRE SUPBEN"]
        self.columnas_color = ["", "TIPO DE PERSONA", "MUNICIPIO", "CULTIVO", "SISTEMA DE RIEGO", "TIPO DE DOCUMENTO QUE ACREDITA EL USO Y APROVECHAMIENTO DE AGUA", "ESTADO DEL PERMISO", "COMPROMISO 1", "COMPROMISO 2", "SOLICITUD", "CURP ", "RECIBO LUZ", "FACTURAS", "ESCRITURAS", "CROQUIS", "RFC ", "BIOMETRICOS", "VERIFICACION"]
        
        self.opciones_icono = ["Punto", "Circulo", "Diamante", "Dona", "Poligono", "Triangulo"]
        self.opciones_tamano = ["Estandar", "Superficie de Riego", "Cuota Energética", "Aprovechamiento", "Consumo Anual", "Consumo por Hectarea", "Volumen Concesionado", "Volumen Consumido", "Consumo entre Volumen", "Supriego entre Supben"]
        self.opciones_color = ["Todos Igual", "Tipo de Persona", "Municipio", "Cultivo", "Sistema de Riego", "Tipo de Permiso", "Vigencia del Permiso", "Compromiso 1", "Compromiso 2", "Checklist_Solicitud", "Checklist_Curp", "Checklist_Recibo_Luz", "Checklist_Facturas", "Checklist_Escrituras", "Checklist_Croquis", "Checklist_RFC", "Checklist_Biometricos", "Checklist_Verificacion"]

        # Etiqueta y entrada para Filtrar por Status
        self.label_status = tk.Label(master, text="1. ¿Filtrar servicios activos?")
        self.label_status.grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        self.status_var = tk.StringVar(value="Sí")
        self.status_entry = ttk.Combobox(master, values=["Sí", "No"], textvariable=self.status_var)
        self.status_entry.grid(row=0, column=1, pady=5, sticky=tk.W)

        # Etiqueta y lista desplegable para Tamaño de la Burbuja
        self.label_icono = tk.Label(master, text="2. Tipo de icono:")
        self.label_icono.grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        self.icono_var = tk.StringVar(value="Punto")
        self.icono_entry = ttk.Combobox(master, values=self.opciones_icono, textvariable=self.icono_var)
        self.icono_entry.grid(row=1, column=1, pady=5, sticky=tk.W)

        # Etiqueta y lista desplegable para Tamaño de la Burbuja
        self.label_tamano = tk.Label(master, text="3. Tamaño del icono:")
        self.label_tamano.grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        self.tamano_var = tk.StringVar(value="Estandar")
        self.tamano_entry = ttk.Combobox(master, values=self.opciones_tamano, textvariable=self.tamano_var)
        self.tamano_entry.grid(row=2, column=1, pady=5, sticky=tk.W)

        # Etiqueta y lista desplegable para Colorear por
        self.label_colorear = tk.Label(master, text="4. Colorear de acuerdo a:")
        self.label_colorear.grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        self.colorear_var = tk.StringVar(value="Todos Igual")
        self.colorear_entry = ttk.Combobox(master, values=self.opciones_color, textvariable=self.colorear_var)
        self.colorear_entry.grid(row=3, column=1, pady=5, sticky=tk.W)

        # Botón para aplicar los filtros
        self.boton_aplicar = tk.Button(master, text="Generar Mapa", command=self.aplicar_filtro)
        self.boton_aplicar.grid(row=4, column=0, columnspan=2, pady=10)

    def aplicar_filtro(self):
        # Obtener las opciones seleccionadas
        status_filtrar = self.status_var.get()
        icono = self.icono_var.get()
        tamano_burbuja = self.tamano_var.get()
        colorear_por = self.colorear_var.get()

        # Obtener los valores correspondientes de las listas
        icono_valor = self.directorio_iconos[self.opciones_icono.index(icono)]
        tamano_burbuja_valor = self.columnas_tamano[self.opciones_tamano.index(tamano_burbuja)]
        colorear_por_valor = self.columnas_color[self.opciones_color.index(colorear_por)]

        # Almacenar las opciones en la lista
        self.opciones_seleccionadas = [status_filtrar, icono_valor, tamano_burbuja_valor, colorear_por_valor,tamano_burbuja,colorear_por]

        # Cerrar la ventana después de aplicar el filtro
        self.master.destroy()

    def iniciar_interfaz(self):
        # Inicia el bucle de la interfaz gráfica y espera hasta que la ventana sea cerrada
        self.master.mainloop()

    def get_opciones_seleccionadas(self):
        return self.opciones_seleccionadas