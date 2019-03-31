#pip install openpyxl
#python -m pip install --upgrade pip            #Actualizar version de pip
#pip install pillow                             #To be able to include images (jpeg, png, bmp) into an openpyxl file.
#pip install xlrd
#pip install pathlib
#pip install xlwt
#pip install pillow
#pip install matplotlib
#pip install pandas
#pip install pypiwin32

import os
import sys
import tkinter
import shutil
import sys
import tkinter as tk
import xml.etree.ElementTree as ET
import difflib
import xlrd
import pandas as pd
import win32com.client
import xml.etree.ElementTree as e

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

CounterFilasExcel=0

estadoCheckButton = 0

################
#### Eventos ###
################
def newProgram():

	global rutaCNT
	global archivoCNTCargado
	global rutaReportePrevio
	global archivoReportePrevioCargado
	global estadoCheckButton

	#Ocultar botones innecesarios.
	button_PreviousReport.grid_remove()
	enable_button.grid_remove()
	button_GenerateReport.grid_remove()
	button_Log.grid_remove()

	button_CNT.configure(style='button_style1.TButton')
	enable_button.deselect()
	button_PreviousReport.configure(style='button_style2.TButton')
	button_GenerateReport.configure(style='button_style2.TButton')
	button_Log.configure(style='button_style2.TButton')

	#Comenzar con el boton de reporte previo deshabilitado.
	estadoCheckButton = 0
	button_PreviousReport.state(["disabled"])

	rutaCNT = "0"
	archivoCNTCargado = 0
	rutaReportePrevio = "0"
	archivoReportePrevioCargado = 0

def exitProgram():
	#Preguntar al usuario si desea salir del programa.
	salir = messagebox.askyesno(message="Do you want to close the program?", title="Close program")
	if(salir == 1):
		sys.exit(0)

def aboutProgram():
	messagebox.showinfo("About EEPROM report generator", "This software has been released by Ruben Barajas Curiel and Julio Cesar Almada Fuerte")

def helpProgram():
	messagebox.showinfo("Help", "Visit the following link to get more information about this software")

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
        if(rutaReportePrevio.find(".xlsx") == -1):        #No se selecciono un archivo XLXS
                archivoReportePrevioCargado = 0
        elif(rutaReportePrevio== ""):                    #Se dio al boton cancelar.
                archivoReportePrevioCargado = 0
        else:                                   #Se selecciono el archivo correctamente.
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
        		
        		#Preguntar directorio para guardar el archivo generado.
                folder_selected = filedialog.askdirectory()

                if(folder_selected != ""):		#Si no se le dio al boton cancelar
                		archivoCNTCargado = 0

                		#Extraer BBNumber y Baseline.
                		BBNumber = rutaCNT.split('_')[3]
                		Baseline = rutaCNT.split('_')[4]

                		#Crear archivo Excel.
                		shutil.copy("EEPROM_Container_Review_Template.xlsx", folder_selected + "/fillexcel.xlsx")

                		#Rellena Excel
                		fillExcel()

                		messagebox.showinfo("Report created", "Report created successfully")

                		button_Log.grid()	#Ahora se puede mostrar el boton del log.
                		button_CNT.configure(style='button_style1.TButton')
                		enable_button.grid_remove()											#Esconder boton para habilitar reporte previo.
                		button_PreviousReport.grid_remove()									#Esconder boton de reporte previo.
                		button_GenerateReport.grid_remove()									#Esconder boton de generar reporte.
                		button_GenerateReport.configure(style='button_style2.TButton')
                
                #Acomodar el archivo para leerlo.
                #previousWorkbook = xlrd.open_workbook(rutaReportePrevio)
                #previousWorksheet = previousWorkbook.sheet_by_name('Checklist')
                #e = xml.etree.ElementTree.parse(rutaCNT).getroot()
                #continuarComparacion = previousWorksheet.nrows
                #contadorCelda = 11			#Posicion del primer elemento NVM data item
                #valorCelda = ""

                #Crear archivo Excel.
                #shutil.copy("EEPROM_Container_Review_Template.xlsx", folder_selected + "/EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + BBNumber + ".xlsx")
                		
                #Mientras existan NVM data item.
                #while(contadorCelda < continuarComparacion):
             	#	#Leer valor de la celda.
                #	valorCelda = previousWorksheet.cell(contadorCelda, 0).value
                #	contadorCelda = contadorCelda + 1

                	#Parser.
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
menubar = Menu(panelImage)

button_style1 = ttk.Style()
button_style2 = ttk.Style()
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
        button_style1.configure("button_style1.TButton", width = 20, padding=5, font=('Helvetica', 10, 'bold'), background = "black", foreground = 'green')
        button_style2.configure("button_style2.TButton", width = 20, padding=5, font=('Helvetica', 10, 'bold'))
        
        button_CNT.configure(style='button_style1.TButton')
        button_PreviousReport.configure(style='button_style2.TButton')
        button_GenerateReport.configure(style='button_style2.TButton')
        button_Log.configure(style='button_style2.TButton')

        #Comenzar con el boton de reporte previo deshabilitado.
        estadoCheckButton = 0
        button_PreviousReport.state(["disabled"])

        root.config(menu=menubar)
        toolBar_Init = Menu(menubar)
        toolBar_About = Menu(menubar)
        toolBar_Init.add_command(label="New", command=newProgram)
        toolBar_Init.add_command(label="Exit", command=exitProgram)
        toolBar_About.add_command(label="About", command=aboutProgram)
        toolBar_About.add_command(label="Help", command=helpProgram)
        menubar.add_cascade(label="File", menu=toolBar_Init)
        menubar.add_cascade(label="Program", menu=toolBar_About)

        #Comenzar proceso.
        root.mainloop()

def fillExcel():
        
        #Carga el archivo Excel anteriormente generado
        wb = load_workbook(filename = folder_selected + "/fillexcel.xlsx")
        ws=wb.active
        
        #Lee archivo XML
        tree = ET.parse(rutaCNT)

        #Obtiene el root del XML
        root = tree.getroot()
        
        #Counter para ir agregando elementos en excel
        CounterFilasExcel=11

        #Busca el nodo sesion en todo el arbol
        for session in root.iter('SESSION'):
                
                #Busca en los tipos de sesiones que nombre tiene
                sessionN=  session.find('SESSION-NAME')

                #Cuando la sesion es ALL                
                if(sessionN.text =='__ALL__'):
                        #Para no alterar el el orden de las filas del excel
                        tempCounter=CounterFilasExcel
                        #Obtiene el Datapointer-name del item
                        for DPN in session.iter('DATAPOINTER-NAME'):
                                tempCounter+=1
                                #guarda el valor en el excel
                                ws['A'+str(tempCounter)]=DPN.text
                        tempCounter=CounterFilasExcel
                        #Obtiene el Datapointer-ident  del item
                        for DPID in session.iter('DATAPOINTER-IDENT'):
                                tempCounter+=1
                                ws['D'+str(tempCounter)]=DPID.text
                        tempCounter=CounterFilasExcel
                        #Obtiene el Datapointer-identifier del item
                        for DFID in session.iter('DATAFORMAT-IDENTIFIER'):
                                #Aqui se aumenta el CounterFilasExcel para que se respeten las filas
                                CounterFilasExcel+=1
                                ws['O'+str(CounterFilasExcel)]=DFID.text

                #Cuando la sesion es Reprog
                if(sessionN.text=='Reprog'):
                        tempCounter=CounterFilasExcel
                        for DPN in session.iter('DATAPOINTER-NAME'):
                                tempCounter+=1
                                ws['A'+str(tempCounter)]=DPN.text
                        tempCounter=CounterFilasExcel
                        for DPID in session.iter('DATAPOINTER-IDENT'):
                                tempCounter+=1
                                ws['D'+str(tempCounter)]=DPID.text
                        tempCounter=CounterFilasExcel
                        for DFID in session.iter('DATAFORMAT-IDENTIFIER'):
                                CounterFilasExcel+=1
                                ws['O'+str(CounterFilasExcel)]=DFID.text
                                #Marca que el use case de que es Reprog
                                ws['K'+str(CounterFilasExcel)]="X"
                
                #Cuando la sesion es DeliveryState
                if(sessionN.text=='DeliveryState'):
                        tempCounter=CounterFilasExcel
                        for DPN in session.iter('DATAPOINTER-NAME'):
                                tempCounter+=1
                                ws['A'+str(tempCounter)]=DPN.text
                        tempCounter=CounterFilasExcel
                        for DPID in session.iter('DATAPOINTER-IDENT'):
                                tempCounter+=1
                                ws['D'+str(tempCounter)]=DPID.text
                        tempCounter=CounterFilasExcel
                        for DFID in session.iter('DATAFORMAT-IDENTIFIER'):
                                CounterFilasExcel+=1
                                ws['O'+str(CounterFilasExcel)]=DFID.text
                                #Marca que el use case de que es DeliveryState
                                ws['I'+str(CounterFilasExcel)]="X"
                
                #Cuando es la sesion es ResetToDeliveryState
                if(sessionN.text=='ResetToDeliveryState'):
                        tempCounter=CounterFilasExcel
                        for DPN in session.iter('DATAPOINTER-NAME'):
                                tempCounter+=1
                                ws['A'+str(tempCounter)]=DPN.text
                        tempCounter=CounterFilasExcel
                        for DPID in session.iter('DATAPOINTER-IDENT'):
                                tempCounter+=1
                                ws['D'+str(tempCounter)]=DPID.text
                        tempCounter=CounterFilasExcel
                        for DFID in session.iter('DATAFORMAT-IDENTIFIER'):
                                CounterFilasExcel+=1
                                ws['O'+str(CounterFilasExcel)]=DFID.text
                                #Marca que el use case de que es ReturnToDeliveryState
                                ws['J'+str(CounterFilasExcel)]="X"
        
        for datablock in root.iter('DATABLOCK'):
                #Busca en los datablock que nombre tiene
                DBN=  datablock.find('DATABLOCK-NAME')
                #Para recorrer el excel
                for i in range(12,CounterFilasExcel):
                        #Compara el valor que tiene la celda de excel con el datablock name
                        if(DBN.text==ws['A'+str(i)].value):
                                j=0
                                #Hay varios DATA por datablock, pero el que se ocupa es el 6to
                                for DPN in datablock.iter('DATA'):
                                        j+=1
                                        if(j==6):
                                                #Se copia la descripcion a la columna comment
                                                ws['Q'+str(i)]=DPN.text  

        #Para checar si se repite algun NVM Item
        for i in range(12,CounterFilasExcel):
                #Agarra cada fila y las comapara con las demas
                temp = ws['A'+str(i)]
                k=i+1
                for j in range(k,CounterFilasExcel):
                        #Aqui se agarra el siguiente en la fila y se checa cada elemento siguiente
                        temp2 = ws['A'+str(j)]
                        #Si son iguales
                        if(temp.value == temp2.value):
                                #Checa el USE CASES de cada uno
                                temp3 = ws['I'+str(j)]
                                if(temp3.value=="X"):
                                        #Marca el use case del que se repite 
                                        ws['I'+str(i)]="X"
                                        #Borra la fila que se repite
                                        ws.delete_rows(j,1)
                                temp3 = ws['J'+str(j)]
                                if(temp3.value=="X"): 
                                        ws['J'+str(i)]="X"
                                        ws.delete_rows(j,1)
                                temp3 = ws['K'+str(j)]
                                if(temp3.value=="X"): 
                                        ws['K'+str(i)]="X"
                                        ws.delete_rows(j,1)
        
        #Guarda los cambios
        wb.save(folder_selected + "/fillexcel.xlsx")        
        #Para ordenar por ID number
        #Selecciona el archivo excel
        excel_file = folder_selected + "/fillexcel.xlsx"
        #lee el archivo
        movies = pd.read_excel(excel_file, skiprows=10)
        #los ordena por numero de ID
        sorted_by_number = movies.sort_values(by='ID number',ascending=True)
        #lo guarda
        sorted_by_number.to_excel(excel_file,index=False)

        #Se crea una copia del Template para poder copiar los datos ordenados al archivo que se generara al final
        shutil.copy("EEPROM_Container_Review_Template.xlsx", folder_selected + "/EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + BBNumber + ".xlsx")
        
        #Carga el archivo Excel anteriormente generado
        wb1 = load_workbook(filename = folder_selected +  "/EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + BBNumber + ".xlsx")
        ws1=wb1.active

        #Carga el archivo con los datos ordenados
        wb=load_workbook(filename = folder_selected +  "/fillexcel.xlsx")
        ws=wb.active
        #Desde aqui agarra los datos
        j=2
        #Los datos los pega ordenados en el archivo excel que es copia del template
        for i in range(12,CounterFilasExcel):
                ws1['A'+str(i)]=ws['A'+str(j)].value
                ws1['D'+str(i)]=ws['D'+str(j)].value
                ws1['I'+str(i)]=ws['I'+str(j)].value
                ws1['J'+str(i)]=ws['J'+str(j)].value
                ws1['K'+str(i)]=ws['K'+str(j)].value
                ws1['O'+str(i)]=ws['O'+str(j)].value
                ws1['Q'+str(i)]=ws['Q'+str(j)].value
                j+=1

        #Asigna los valores de BBNumber y Baseline a sus respectivas celdas
        ws1['D3']=BBNumber
        ws1['D4']=Baseline

        #Guarda el archivo
        wb1.save(folder_selected +  "/EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + BBNumber + ".xlsx")
        #Borra el archivo que tiene los datos ordenados
        os.remove(folder_selected + "/fillexcel.xlsx")
        if(archivoReportePrevioCargado):
                MergeExcel()

def MergeExcel():
        print("hola")

def main():
		ventana()

#################################
if __name__== "__main__":
		main()
