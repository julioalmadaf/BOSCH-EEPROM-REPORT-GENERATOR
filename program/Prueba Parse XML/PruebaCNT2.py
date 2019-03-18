import xml.etree.ElementTree as ET

RC=0
DSC=0
RDSC=0

tree = ET.parse('EEPROM_Container_BB95650_BSS2007_V3_IPBCSWNonXCP.cnt')

root = tree.getroot()

print(root.tag, root.attrib)
print('Ahi te van los hijos')

for child in root:
    print (child.tag, child.attrib)

reprog = root[1][0][1][0].text
if (reprog == "Reprog"):
    print("SESSION-NAME = Reprog")
reprogSig = root[1][0][1][2]
for DName in reprogSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)
    RC+=1
print(RC)     

DS = root[1][0][2][0].text
if (DS == "DeliveryState"):
    print("SESSION-NAME = DeliveryState")
DSSig = root[1][0][2][2]
for DName in DSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)
    DSC+=1
print(DSC)

RDS = root[1][0][3][0].text
if (RDS == "ResetToDeliveryState"):
    print("SESSION-NAME = ResetToDeliveryState")
RDSSig = root[1][0][3][2]
for DName in RDSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)
    RDSC+=1
print(RDSC)


#Datapointer-name de la sesion Reprog
#Dummy_testmoduleCnt
#print(root[1][0][1][2][5][1].tag)
#print(root[1][0][1][2][5][1].text)
