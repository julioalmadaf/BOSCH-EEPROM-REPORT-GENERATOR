**EEPROM Report Generator**

**Overview**

One of the reports required for releasing the SW is the EEPROM review report; this document is created using the EEPROM container, which is a special .cnt file which is based on XML structure. The .cnt file has all the EEPROM elements that will be saved in the non-volatile memory; the report is used by the component responsible to evaluate the values and the uses cases of each of the elements in the EEPROM memory.

This tool is intended to automate the report generation and improve the release time. 
A program for Windows 10 that takes the information from a container to an Excel file automatically.
This was developed to make fast and efficient reports that required hours to do manually and now they can be done in just a couple of seconds.

For this project there are 2 inputs:
1) A container(A XML file)
2) A previous report(Optional).

When inputting those files through the UI, the python program  will read, process and generate a new final report and a Log of actions.

**Authors**
- Julio Almada
- Ruben Barajas
