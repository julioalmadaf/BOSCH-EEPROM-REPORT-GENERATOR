import xml.etree.ElementTree as ET
#Counter de datapointers para cada sesion
RC=0
DSC=0
RDSC=0
#Lee archivo XML
tree = ET.parse('EEPROM_Container_BB95650_BSS2007_V3_IPBCSWNonXCP.cnt')

f=open("logfiletest.txt","w+")
f.write("\r\nExcel file being processed\r\n\r\n")

#Obtiene el root
root = tree.getroot()
f.write("root: " + str(root) + " added to logtestfile.txt.\r\n\r\n")

for session in root.iter('SESSION'):
        f.write("session: " + str(session) + " added to logtestfile.txt.\r\n\r\n")
        sessionN=  session.find('SESSION-NAME')
        if(sessionN.text=='Reprog'):
                for MEH in session.find('DATAPOINTERS'):
                        RC+=1
                        print(RC)
                for DPN in session.iter('DATAPOINTER-NAME'):
                        f.write("Datapointer name: " + str(DPN) + " added to logtestfile.txt.\r\n\r\n")
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

f.write("END OF LOG FILE\r\n")
f.close()