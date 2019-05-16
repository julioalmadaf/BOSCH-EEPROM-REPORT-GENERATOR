import xml.etree.ElementTree as ET
from openpyxl import Workbook

wb= Workbook()
ws = wb.active

ws['A4'] = 4

wb.save('Excelvalues.xlsx')
                    