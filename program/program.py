#pip install openpyxl
#python -m pip install --upgrade pip            #Actualizar version de pip
#pip install pillow                             #To be able to include images (jpeg, png, bmp) into an openpyxl file.
#pip install xlrd
#pip install pathlib

#pip install pillow
#pip install matplotlib

import os
import sys
import tkinter
import shutil
import sys
import tkinter as tk

from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
from PIL import ImageTk, Image
from pathlib import Path
from os import listdir

###########################
#### Variables globales ###
###########################
rutaCNT = "0"
archivoCNTCargado = 0
rutaReportePrevio = "0"
archivoReportePrevioCargado = 0
estadoCheckButton = 0

################
#### Eventos ###
################
def cntButton():

        global rutaCNT
        global archivoCNTCargado

        #Abrir Dialog Box para buscar el archivo.
        label.configure(text="Opening CNT file")
        rutaCNT = filedialog.askopenfilename(filetypes = (("All CNT files","*.cnt"),("All files","*.*")))

        #Verificar que sea archivo CNT o que se haya agregado un archivo.
        if(rutaCNT.find(".cnt") == -1):         #No se selecciono un archivo CNT
                archivoCNTCargado = 0
        elif(rutaCNT == ""):                    #Se dio al boton cancelar.
                archivoCNTCargado = 0
        else:                                   #Se selecciono el archivo correctamente.
                archivoCNTCargado = 1
                button_CNT.configure(style='button_style2.TButton')
                enable_button.grid()			#Ahora se puede mostrar el checkbuton para habilitar el boton de reporte previo.
                button_PreviousReport.grid()	#Ahora se puede mostrar el boton de reporte previo como opcional.
                button_GenerateReport.grid()	#Ahora se puede mostrar el boton de generar reporte.
                button_GenerateReport.configure(style='button_style1.TButton')
                button_Log.grid_remove()		#Esconder boton de log.

def enableButtonRP():

	global estadoCheckButton

	estadoCheckButton = estadoCheckButton ^ 1

	#Cambiar estado del boton de reporte previo dependiendo del CheckButton.
	if(estadoCheckButton == 0):
		button_PreviousReport.state(["disabled"])
	else:
		button_PreviousReport.state(["!disabled"])

def previousReport():

        global rutaReportePrevio
        global archivoReportePrevioCargado

        #Abrir Dialog Box para buscar el archivo.
        label.configure(text="Opening previous report file")
        rutaReportePrevio = filedialog.askopenfilename(filetypes = (("All Excel files","*.xlsx"),("All files","*.*")))

        #Verificar que sea archivo XLSX o que se haya agregado un archivo.
        if(rutaCNT.find(".xlsx") == -1):         #No se selecciono un archivo XLXS
                archivoReportePrevioCargado = 0
        elif(rutaCNT == ""):                    #Se dio al boton cancelar.
                archivoReportePrevioCargado = 0
        else:                                   #Se selecciono el archivo correctamente.
                archivoReportePrevioCargado = 1

def createReport():

        global archivoCNTCargado
        global archivoReportePrevioCargado
        global rutaCNT
        global rutaReportePrevio

        label.configure(text="Creating report")

        #Garantizar que se haya seleccionado un archivo CNT.
        if(archivoCNTCargado == 1):
        		
        		#Preguntar directorio para guardar el archivo generado.
                folder_selected = filedialog.askdirectory()

                if(folder_selected != ""):		#Si no se le dio al boton cancelar
                		archivoCNTCargado = 0

                		#Extraer BBNumber y Baseline.
                		BBNumber = rutaCNT.split('_')[3]
                		Baseline = rutaCNT.split('_')[4]

                		#Crear archivo Excel.
                		shutil.copy("EEPROM_Container_Review_Template.xlsx", folder_selected + "/EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + BBNumber + ".xlsx")

                		messagebox.showinfo("Report created", "Report created successfully")

                		button_Log.grid()	#Ahora se puede mostrar el boton del log.
                		button_CNT.configure(style='button_style1.TButton')
                		enable_button.grid_remove()				#Esconder boton para habilitar reporte previo.
                		button_PreviousReport.grid_remove()		#Esconder boton de reporte previo.
                		button_GenerateReport.grid_remove()		#Esconder boton de generar reporte.
                		button_GenerateReport.configure(style='button_style2.TButton')

        else:
                messagebox.showerror("Error", "Not .cnt file selected")

        label.configure(text="EEPROM report generator")

def verifyLog():
        label.configure(text="Log")
        button_Log.grid_remove()		#Esconder boton de log.

################
#### Objetos ###
################
root = Tk()

rutaActual = os.getcwd()
img = ImageTk.PhotoImage(Image.open(rutaActual + "/bosch.png"))

panelElements = ttk.Frame(root, padding=(3,3,12,12))
panelImage = ttk.Frame(panelElements, borderwidth=5, relief="sunken", width=200, height=200)

label = ttk.Label(panelElements, text="Bosch EEPROM generator", font=("Tahoma", 25, 'bold'))
button_CNT = ttk.Button(panelElements, text="Select CNT file", style="TButton", command=cntButton)
button_PreviousReport = ttk.Button(panelElements, text="Select previous report", style="TButton", command=previousReport)
button_GenerateReport = ttk.Button(panelElements, text="Generate report", style="TButton", command=createReport)
button_Log = ttk.Button(panelElements, text="LOG", style="TButton", command=verifyLog)
enable_button = Checkbutton(panelElements, text="Enable previous report button", onvalue=1,offvalue=0, command=enableButtonRP)
image = tk.Label(panelImage, image = img)

##################
#### Funciones ###
##################

def ventana():

        #Titulo de la ventana.
        root.title("Bosch") 

        #Configurar paneles.
        panelElements.grid(column=0, row=0, sticky=(N, S, E, W))
        panelImage.grid(column=0, row=1, columnspan=2, rowspan=7, sticky=(N, S, E, W))

        #Configurar elementos (botones, etiqueta, imagen, etc).
        image.pack(side = "bottom", fill = "both", expand = "yes")
        label.grid(column=0, row=0, columnspan=4, sticky=(N, W))
        button_CNT.grid(column=3, row=3)
        enable_button.grid(column=3, row=4)
        button_PreviousReport.grid(column=3, row=5)
        button_GenerateReport.grid(column=3, row=6)
        button_Log.grid(column=3, row=7)
        
        #Ocultar botones innecesarios.
        button_PreviousReport.grid_remove()
        enable_button.grid_remove()
        button_GenerateReport.grid_remove()
        button_Log.grid_remove()

        #Fondos y colores.
        style = ttk.Style(root)
        style.configure('TLabel', background='white')	#Background y foreground de la etiqueta.
        style.configure('TFrame', background='white')	#Background y foreground del Frame.

        #Estilo de los elementos.
        button_style1 = ttk.Style()
        button_style2 = ttk.Style()
        button_style1.configure("button_style1.TButton", width = 20, padding=5, font=('Helvetica', 10, 'bold'), background = "black", foreground = 'green')
        button_style2.configure("button_style2.TButton", width = 20, padding=5, font=('Helvetica', 10, 'bold'))
        
        button_CNT.configure(style='button_style1.TButton')
        button_PreviousReport.configure(style='button_style2.TButton')
        button_GenerateReport.configure(style='button_style2.TButton')
        button_Log.configure(style='button_style2.TButton')

        #Comenzar con el boton de reporte previo deshabilitado.
        estadoCheckButton = 0
        button_PreviousReport.state(["disabled"])

        #Comenzar proceso.
        root.mainloop()

def main():
		ventana()

#################################
if __name__== "__main__":
		main()