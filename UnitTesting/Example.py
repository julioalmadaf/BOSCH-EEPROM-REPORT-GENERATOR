import xml.etree.ElementTree as ET
tree = ET.parse('Example.xml')
root = tree.getroot()
print(root.tag, root.attrib)                #Nombre del root
for child in root:
    print (child.tag, child.attrib)         #Tag de cada Child
print(root[1][0].text)                      #PARA los childs anidados
