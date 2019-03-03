#pip install openpyxl
#python -m pip install --upgrade pip		#Actualizar version de pip
#pip install pillow							#To be able to include images (jpeg, png, bmp) into an openpyxl file.
#pip install xlrd

#pip install pathlib

import os
import sys
from pathlib import Path
from os import listdir
 
def list_files(directory, extension):
    return (f for f in listdir(directory) if f.endswith('.' + extension))

def main():

	#Obtener el path actual.
	mypath = Path().absolute()

	#Buscar archivos .cnt
	files = list_files(mypath,"cnt")
	
	#De cada archivo cnt encontrado, segmentarlo para encontrar el BBNumber y el Baseline.
	for f in files:
		BBNumber = f.split('_')[2]
		Baseline = f.split('_')[3]
		
	print(BBNumber)
	print(Baseline)
	

#################################
if __name__== "__main__":
	main()