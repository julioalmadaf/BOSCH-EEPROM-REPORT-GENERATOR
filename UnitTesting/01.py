import os
import sys
import tkinter as tk
import xml.etree.ElementTree as ET
import xml.etree.ElementTree as e
import xlrd
import xlwt
import shutil
import pandas as pd
import win32com.client

from os import listdir
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
from pathlib import Path
from PIL import ImageTk, Image
from openpyxl import load_workbook

#rutaCNT = "C:/Users/Barajas/Desktop/ExternalFiles/EEPROM_Container_BB95650_BSS2007_V3_IPBCSWNonXCP.cnt"
rutaCNT = "C:/Users/Barajas/Desktop/ExternalFiles/EEPROM_Container_IPBCSWNonXCP.cnt"
#rutaCNT = "C:/Users/Barajas/Desktop/ExternalFiles/EEPROM_Container_BB95650_BSS2007_V3_IPBCSWNonXCP.cnp"

BBNumber = ""
Baseline = ""
if(os.path.isfile(rutaCNT) == True):

	encontrarBBN = rutaCNT.find("BB")
	if(encontrarBBN != -1):
		BBNumber = rutaCNT.split('_')[2]
		print(BBNumber)
	else:
		print("BBNumber not found in the .cnt name\r\n")

	encontrarBSS = rutaCNT.find("BSS")
	if(encontrarBSS != -1):
		Baseline = rutaCNT.split('_')[3]
		print(Baseline)
	else:
		print("Baseline not found in the .cnt name\r\n")

	#Crear archivo Excel.
	rutaArchivoCNT = os.path.dirname(rutaCNT)
	shutil.copy("EEPROM_Container_Review_Template.xlsx", rutaArchivoCNT + "/fillexcel.xlsx")

	#Carga el archivo Excel anteriormente generado.
	wb = load_workbook(filename = rutaArchivoCNT + "/fillexcel.xlsx")
	ws = wb.active

	#Lee archivo XML.
	tree = ET.parse(rutaCNT)

	#Obtiene el root del XML.
	root = tree.getroot()
else:
	print("File not found\r\n")
