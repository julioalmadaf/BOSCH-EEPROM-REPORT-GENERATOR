#pip install openpyxl
#python -m pip install --upgrade pip		#Actualizar version de pip
#pip install pillow							#To be able to include images (jpeg, png, bmp) into an openpyxl file.
#pip install xlrd
#pip install pathlib

import os
import sys
import tkinter
import shutil
import sys 

from tkinter import messagebox
from pathlib import Path
from os import listdir

import tkinter as tk
from tkinter import filedialog

def list_files(directory, extension):
    return (f for f in listdir(directory) if f.endswith('.' + extension))

def main():

	#Seleccionar container.
	buscarArchivoCNT = 0
	while buscarArchivoCNT == 0:
		
		#Abrir Dialog Box y seleccionar archivo.
		root = tk.Tk()
		root.withdraw()
		file_path = filedialog.askopenfilename()
		print(file_path)

		#Validar que sea un archivo .cnt
		if(file_path.find(".cnt") != -1):
			buscarArchivoCNT = 1
		else:
			messagebox.showerror("Error", "Not .cnt file selected")
	
	#Extraer BBNumber y Baseline.
	BBNumber = file_path.split('_')[3]
	Baseline = file_path.split('_')[4]
	
	#Crear archivo Excel.
	shutil.copy("EEPROM_Container_Review_Template.xlsx", "EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_" + BBNumber + ".xlsx")
	

#################################
if __name__== "__main__":
	main()