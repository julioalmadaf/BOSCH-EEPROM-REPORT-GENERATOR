#pip install pandas

#Libreria de pandas
import pandas as pd

#Libreria para abrir archivo
from tkinter import filedialog

#Selecciona el archivo excel
excel_file = filedialog.askopenfilename(filetypes = (("All xls files","*.xls"),("All files","*.*")))

#lee el archivo
movies = pd.read_excel(excel_file)

#imprime lo que hay en el excel
print(movies)

#los ordena por numero
sorted_by_number = movies.sort_values(by='Year',ascending=True)

#lo imprime
print(sorted_by_number)

#lo guarda
sorted_by_number.to_excel(excel_file,index=False)