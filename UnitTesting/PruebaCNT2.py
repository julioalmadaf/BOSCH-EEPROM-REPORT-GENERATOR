import xml.etree.ElementTree as ET

RC=0
DSC=0
RDSC=0

tree = ET.parse('EEPROM_Container_BB95650_BSS2007_V3_IPBCSWNonXCP.cnt')

root = tree.getroot()

print(root.tag, root.attrib)

#"root" de Reprog
reprog = root[1][0][1][0].text
#checa si coincide con la sesion Reprog
if (reprog == "Reprog"):
    print("SESSION-NAME = Reprog")
reprogSig = root[1][0][1][2]
for DName in reprogSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)
    #Cuenta cuantos datapointers hay dentro de reprog
    RC+=1
print(RC)

#"root" de deliverystate
DS = root[1][0][2][0].text
#checa si coincide con la sesion DeliveryState
if (DS == "DeliveryState"):
    print("SESSION-NAME = DeliveryState")
DSSig = root[1][0][2][2]
for DName in DSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)
    #Cuenta cuantos datapointers hay dentro de DeliveryState
    DSC+=1
print(DSC)

#"root" de return to delivery state
RDS = root[1][0][3][0].text
#checa si coincide con la sesion ResetToDeliveryState
if (RDS == "ResetToDeliveryState"):
    print("SESSION-NAME = ResetToDeliveryState")
RDSSig = root[1][0][3][2]
for DName in RDSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    print(DName.tag, DName.text)
    #Cuenta cuantos datapointers hay dentro de ResetToDeliveryState
    RDSC+=1
print(RDSC)


#Datapointer-name de la sesion Reprog
#Dummy_testmoduleCnt
#ECU-MEM [MEM][SESSIONS][SESSION][DATAPOINTERS][DATAPOINTER][DATAPOINTER-NAME]
print(root[1][0][1][2][4][0].tag)
print(root[1][0][1][2][4][0].text)

jeje = root[1][0][1][2][4][0].text
print(jeje)