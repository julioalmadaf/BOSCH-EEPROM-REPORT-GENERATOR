import xml.etree.ElementTree as ET
#Counter de datapointers para cada sesion
RC=0
DSC=0
RDSC=0
#Lee archivo XML
tree = ET.parse('EEPROM_Container_BB95650_BSS2007_V3_IPBCSWNonXCP.cnt')
#Obtiene el root
root = tree.getroot()

for session in root.iter('SESSION'):
        sessionN=  session.find('SESSION-NAME')
        if(sessionN.text=='Reprog'):
                for MEH in session.find('DATAPOINTERS'):
                        RC+=1
                        print(RC)
                for DPN in session.iter('DATAPOINTER-NAME'):
                        print(DPN.text)
                for DPID in session.iter('DATAPOINTER-IDENT'):
                        print(DPID.text)
                for DFID in session.iter('DATAFORMAT-IDENTIFIER'):
                        print(DFID.text)
        if(sessionN.text=='DeliveryState'):
                for MEH in session.find('DATAPOINTERS'):
                        DSC+=1
                        print(DSC)
                for DPN in session.iter('DATAPOINTER-NAME'):
                        print(DPN.text)
                for DPID in session.iter('DATAPOINTER-IDENT'):
                        print(DPID.text)
                for DFID in session.iter('DATAFORMAT-IDENTIFIER'):
                        print(DFID.text)
        if(sessionN.text=='ResetToDeliveryState'):
                for MEH in session.find('DATAPOINTERS'):
                        RDSC+=1
                        print(RDSC)
                for DPN in session.iter('DATAPOINTER-NAME'):
                        print(DPN.text)
                for DPID in session.iter('DATAPOINTER-IDENT'):
                        print(DPID.text)
                for DFID in session.iter('DATAFORMAT-IDENTIFIER'):
                        print(DFID.text)
                        


'''
#Reprog
#reprogSig = root[1][0][1][2]
for Lreprog in reprogSig.iter('DATAPOINTER-NAME'):
    #Cuenta cuantos datapointers hay dentro de reprog
    RC+=1
print(RC)

#Imprime cada dato importante de cada datapointer
for i in range(0,RC):
    for j in range (0,4):
        print(root[1][0][1][2][i][j].tag)
        print(root[1][0][1][2][i][j].text)


#cheque como moverme entre los datos necesarios de cada datapointer
print(root[1][0][1][2][0][0].tag)
print(root[1][0][1][2][0][0].text)
print(root[1][0][1][2][0][1].tag)
print(root[1][0][1][2][0][1].text)
print(root[1][0][1][2][0][3].tag)
print(root[1][0][1][2][0][3].text)

#Guarda lo necesario para el excel
DNAME = root[1][0][1][2][0][0].text
DID = root[1][0][1][2][0][1].text
DFORMAT= root[1][0][1][2][0][3].text
#los imprime
print(DNAME)
print(DID)
print(DFORMAT)

#Deliverystate
DSSig = root[1][0][2][2]
for DName in DSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    #Cuenta cuantos datapointers hay dentro de DeliveryState
    DSC+=1
print(DSC)

#Imprime cada dato importante de cada datapointer
for i in range(0,DSC):
    for j in range (0,4):
        print(root[1][0][2][2][i][j].tag)
        print(root[1][0][2][2][i][j].text)

#Return to delivery state
RDSSig = root[1][0][3][2]
for DName in RDSSig.findall("./DATAPOINTER/DATAPOINTER-NAME"):
    #Cuenta cuantos datapointers hay dentro de ResetToDeliveryState
    RDSC+=1
print(RDSC)

#Imprime cada dato importante de cada datapointer
for i in range(0,RDSC):
    for j in range (0,4):
        print(root[1][0][3][2][i][j].tag)
        print(root[1][0][3][2][i][j].text)


#Datapointer-name de la sesion Reprog
#Dummy_testmoduleCnt
#ECU-MEM [MEM][SESSIONS][SESSION][DATAPOINTERS][DATAPOINTER][DATAPOINTER-NAME]
print(root[1][0][1][2][4][0].tag)
print(root[1][0][1][2][4][0].text)

jeje = root[1][0][1][2][4][0].text
print(jeje)
'''