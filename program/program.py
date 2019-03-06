#pip install openpyxl
#python -m pip install --upgrade pip            #Actualizar version de pip
#pip install pillow                                                     #To be able to include images (jpeg, png, bmp) into an openpyxl file.
#pip install xlrd
#pip install pathlib

import os
import sys
import tkinter
import shutil
import sys 

from tkinter import messagebox
from tkinter import filedialog
from tkinter import *
from pathlib import Path
from os import listdir

###########################
#### Variables globales ###
###########################
rutaCNT = "0"
archivoCNTCargado = 0
rutaReportePrevio = "0"
archivoReportePrevioCargado = 0

################
#### Eventos ###
################
def cntButton():

        global rutaCNT
        global archivoCNTCargado

        #Abrir Dialog Box para buscar el archivo.
        lbl.configure(text="Opening CNT file")
        rutaCNT = filedialog.askopenfilename(filetypes = (("All CNT files","*.cnt"),("All files","*.*")))

        #Verificar que sea archivo CNT o que se haya agregado un archivo.
        if(rutaCNT.find(".cnt") == -1):         #No se selecciono un archivo CNT
                archivoCNTCargado = 0
        elif(rutaCNT == ""):                    #Se dio al boton cancelar.
                archivoCNTCargado = 0
        else:                                   #se selecciono el archivo correctamente.
                archivoCNTCargado = 1

def previousReport():

        global rutaReportePrevio
        global archivoReportePrevioCargado

        #Abrir Dialog Box para buscar el archivo.
        lbl.configure(text="Opening previous report file")
        rutaReportePrevio = filedialog.askopenfilename(filetypes = (("All Excel files","*.xlsx"),("All files","*.*")))

        #Verificar que sea archivo XLSX o que se haya agregado un archivo.
        if(rutaCNT.find(".xlsx") == -1):         #No se selecciono un archivo XLXS
                archivoReportePrevioCargado = 0
        elif(rutaCNT == ""):                    #Se dio al boton cancelar.
                archivoReportePrevioCargado = 0
        else:                                   #se selecciono el archivo correctamente.
                archivoReportePrevioCargado = 1

def createReport():

        global archivoCNTCargado
        global archivoReportePrevioCargado
        global rutaCNT
        global rutaReportePrevio

        lbl.configure(text="Creating report")

        #Garantizar que se haya seleccionado un archivo CNT.
        if(archivoCNTCargado == 1):
                archivoCNTCargado = 0

                #Extraer BBNumber y Baseline.
                BBNumber = rutaCNT.split('_')[3]
                Baseline = rutaCNT.split('_')[4]

                #Preguntar directorio para guardar el archivo generado.
                folder_selected = filedialog.askdirectory()

                #Crear archivo Excel.
                shutil.copy("EEPROM_Container_Review_Template.xlsx", folder_selected + "/EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + BBNumber + ".xlsx")

                messagebox.showinfo("Report created", "Report created successfully")

        else:
                messagebox.showerror("Error", "Not .cnt file selected")

        lbl.configure(text="EEPROM report generator")

def verifyLog():
        lbl.configure(text="Log")

################
#### Objetos ###
################
window = Tk()
lbl = Label(window, text="EEPROM report generator", font=("Arial Bold", 30))
btn1 = Button(window, text="Open CNT file", command=cntButton)
btn2 = Button(window, text="Open previous report", command=previousReport)
btn3 = Button(window, text="Create report", command=createReport)
btn4 = Button(window, text="Verify log", command=verifyLog)

##################
#### Funciones ###
##################
def list_files(directory, extension):
        return (f for f in listdir(directory) if f.endswith('.' + extension))

def ventana():

        #Titulo de la ventana.
        window.title("Bosch") 

        #Tamaño de la ventana.
        window.geometry('550x350')

        #Etiqueta.
        lbl.grid(column=0, row=0)

        #Botones.
        btn1.grid(column=0, row=1)
        btn2.grid(column=1, row=2)
        btn3.grid(column=0, row=3)
        btn4.grid(column=1, row=4)

        #Comenzar proceso.
        window.mainloop()

def main():

        ventana()

#################################
if __name__== "__main__":
        main()