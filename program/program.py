#pip install openpyxl
#python -m pip install --upgrade pip		#Actualizar version de pip
#pip install pillow							#To be able to include images (jpeg, png, bmp) into an openpyxl file.
#pip install xlrd
#pip install pathlib

import os
import sys
import tkinter
import openpyxl

from openpyxl import Workbook
from tkinter import messagebox
from pathlib import Path
from os import listdir

import tkinter as tk
from tkinter import filedialog
wb = Workbook()
ws = wb.active

def list_files(directory, extension):
    return (f for f in listdir(directory) if f.endswith('.' + extension))

def main():

	#Seleccionar container
	buscarArchivoCNT = 0
	while buscarArchivoCNT == 0:
		
		#Abrir Dialog Box y seleccionar archivo
		root = tk.Tk()
		root.withdraw()
		file_path = filedialog.askopenfilename()
		print(file_path)

		#Validar que sea un archivo .cnt
		if(file_path.find(".cnt") != -1):
			buscarArchivoCNT = 1
		else:
			messagebox.showerror("Error", "Not .cnt file selected")
	
	#
	BBNumber = file_path.split('_')[3]
	Baseline = file_path.split('_')[4]
	wb.save(f"EEPROM_Container_Review_Checkist_GM_iPB_GlobalB_{BBNumber}.xlsx")
	
	print(BBNumber)
	print(Baseline)

#################################
if __name__== "__main__":
	main()