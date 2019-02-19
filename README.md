#Rationale
Formal software release processes and automotive industry standards require that all the official software releases should be traceable and well documented; the current trends in software development include the implementation of continuous integration/continuous testing paradigms into the release process to expedite the delivery of the Software and automate repetitive processes in order to adapt to the new agile methodologies, which demand high adaptability to engineering changes.

#Overview
One of the reports required for releasing the SW is the EEPROM review report; this document is created using the EEPROM container, which is a special .cnt file which is based on XML structure. The .cnt file has all the EEPROM elements that will be saved in the non-volatile memory; the report is used by the component responsible to evaluate the values and the uses cases of each of the elements in the EEPROM memory.
This tool is intended to automate the report generation and improve the release time. 
