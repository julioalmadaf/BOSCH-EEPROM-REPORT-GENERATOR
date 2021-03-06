import xml.etree.ElementTree as ET

tree = ET.parse("C:/Users/Barajas/Desktop/ExternalFiles/EEPROM_Container_BB95650_BSS2007_V3_IPBCSWNonXCP.cnt")

root = tree.getroot()

print(root.tag, root.attrib)
print('Ahi te van los hijos')

for child in root:
    print (child.tag, child.attrib)

#Busca solo por los Session-name en el archivo CNT

'''

'''
#Session-name = Reprog
#print(root[1][0][1][0].tag)
#print(root[1][0][1][0].text)

reprog = root[1][0][1][0].text
if (reprog == "Reprog"):
    print("SESSION-NAME = Reprog")
reprogSig = root[1][0][1][2]
for DName in reprogSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)



'''
#Datapointer-name de la sesion Reprog
#Dummy_testmoduleCnt
print(root[1][0][1][2][0][0].tag)
print(root[1][0][1][2][0][0].text)
'''
'''
#Session-name = DeliveryState
print(root[1][0][2][0].tag)
print(root[1][0][2][0].text)
'''

DS = root[1][0][2][0].text
if (DS == "DeliveryState"):
    print("SESSION-NAME = DeliveryState")
DSSig = root[1][0][2][2]
for DName in DSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)

'''
#Session-name = ReturnToDeliveryState
print(root[1][0][3][0].tag)
print(root[1][0][3][0].text)
'''

RDS = root[1][0][3][0].text
if (RDS == "ResetToDeliveryState"):
    print("SESSION-NAME = ResetToDeliveryState")
RDSSig = root[1][0][3][2]
for DName in RDSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)