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
import xml.etree.ElementTree as ET

from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
from PIL import ImageTk, Image
from pathlib import Path
from os import listdir
from openpyxl import load_workbook

###########################
#### Variables globales ###
###########################
rutaCNT = "0"
archivoCNTCargado = 0
rutaReportePrevio = "0"
archivoReportePrevioCargado = 0
BBNumber = 0
Baseline = 0
folder_selected="0"

#Variables XML
#Lee archivo XML
tree = "0"

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
        else:                                   #se selecciono el archivo correctamente.
                archivoCNTCargado = 1

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
        else:                                   #se selecciono el archivo correctamente.
                archivoReportePrevioCargado = 1

def createReport():

        global archivoCNTCargado
        global archivoReportePrevioCargado
        global rutaCNT
        global rutaReportePrevio
        global folder_selected
        global BBNumber
        global Baseline

        label.configure(text="Creating report")

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
                #Rellena Excel
                fillExcel()

        else:
                messagebox.showerror("Error", "Not .cnt file selected")

        label.configure(text="EEPROM report generator")

def verifyLog():
        label.configure(text="Log")

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
image = tk.Label(panelImage, image = img)

##################
#### Funciones ###
##################

def ventana():

        #Titulo de la ventana.
        root.title("Bosch") 

        #Configurar paneles.
        panelElements.grid(column=0, row=0, sticky=(N, S, E, W))
        panelImage.grid(column=0, row=1, columnspan=2, rowspan=6, sticky=(N, S, E, W))

        #Configurar elementos (botones, etiqueta e imagen).
        image.pack(side = "bottom", fill = "both", expand = "yes")
        label.grid(column=0, row=0, columnspan=4, sticky=(N, W))
        button_CNT.grid(column=3, row=3)
        button_PreviousReport.grid(column=3, row=4)
        button_GenerateReport.grid(column=3, row=5)
        button_Log.grid(column=3, row=6)
        
        #Comenzar proceso.
        root.mainloop()

def fillExcel():
        

        #Carga el archivo Excel anteriormente generado
        wb = load_workbook(filename = folder_selected + "/EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + str(BBNumber) + ".xlsx")
        ws=wb.active
        
        #Lee archivo XML
        tree = ET.parse(rutaCNT)

        #Obtiene el root del XML
        root = tree.getroot()
        
        #Counter de datapointers para cada sesion
        RC=0            #Reprog
        DSC=0           #DeliveryState
        RDSC=0          #ReturnToDeliveryState

        CounterFilasExcel=11

        #Reprog
        reprogSig = root[1][0][1][2]
        for Lreprog in reprogSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
                #Cuenta cuantos datapointers hay dentro de reprog
                RC+=1
        
        
        #Imprime cada dato importante de cada datapointer
        for i in range(0,RC):
                CounterFilasExcel+=1
                ws['K'+str(CounterFilasExcel)]="X"
                for j in range (0,4):
                        if(j==0): ws['A'+str(CounterFilasExcel)]=root[1][0][1][2][i][j].text
                        if(j==1): ws['D'+str(CounterFilasExcel)]=root[1][0][1][2][i][j].text
                        if(j==3): ws['O'+str(CounterFilasExcel)]=root[1][0][1][2][i][j].text
                        
        #Deliverystate
        DSSig = root[1][0][2][2]
        for DName in DSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
                #Cuenta cuantos datapointers hay dentro de DeliveryState
                DSC+=1

        #Imprime cada dato importante de cada datapointer
        for i in range(0,DSC):
                CounterFilasExcel+=1
                ws['I'+str(CounterFilasExcel)]="X"
                for j in range (0,4):
                        if(j==0): ws['A'+str(CounterFilasExcel)]=root[1][0][2][2][i][j].text
                        if(j==1): ws['D'+str(CounterFilasExcel)]=root[1][0][2][2][i][j].text
                        if(j==3): ws['O'+str(CounterFilasExcel)]=root[1][0][2][2][i][j].text

        #Return to delivery state
        RDSSig = root[1][0][3][2]
        for DName in RDSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
                #Cuenta cuantos datapointers hay dentro de ResetToDeliveryState
                RDSC+=1

        #Imprime cada dato importante de cada datapointer
        for i in range(0,RDSC):
                CounterFilasExcel+=1
                ws['J'+str(CounterFilasExcel)]="X"
                for j in range (0,4):
                        if(j==0): ws['A'+str(CounterFilasExcel)]=root[1][0][3][2][i][j].text
                        if(j==1): ws['D'+str(CounterFilasExcel)]=root[1][0][3][2][i][j].text
                        if(j==3): ws['O'+str(CounterFilasExcel)]=root[1][0][3][2][i][j].text

        #Asigna los valores de BBNumber y Baseline a sus respectivas celdas
        ws['D3']=BBNumber
        ws['D4']=Baseline
        
        #Guarda los cambios
        wb.save(folder_selected + "/EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + str(BBNumber) + ".xlsx")        

def main():

        ventana()

#################################
if __name__== "__main__":
        main()