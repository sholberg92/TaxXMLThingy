import csv
import tkinter as tk
from tkinter import filedialog
import xml.etree.ElementTree as ET



root = tk.Tk()
root.withdraw()
filePath = filedialog.askopenfilename(filetypes=(("CSV Files", "*.csv"),),
                                      title="Select CSV file")

taxDocument = open(filePath)

csvTaxDocument = csv.reader(taxDocument)

xmlFilePath = filedialog.askopenfilename(filetypes=(("XML Files", "*.xml"),),
                                         title="Select XML file")
tree = ET.parse(xmlFilePath)

root = tree.getroot()
outputPath = filedialog.asksaveasfilename(filetypes=(("XML Files", "*.xml"),),
                                          title="Save XML files",
                                          defaultextension=".xml")
fnInsertPos = outputPath.find('.xml')
rows = []

for row in csvTaxDocument:
    if row[0] != '':
        rows.append(row)
    
print(rows)
numberOfRows = 14 # set to how many rows in the table in the destination document

print(len(rows))

for x in range(0, len(rows), numberOfRows):
    i = x # preserve x index and use i for inner loop
    for xmlRow in root.find('Page1').find('Table_Line1'):       
        j = 0 # index for iterating over rows in XML doc             
        for xmlCol in xmlRow:
            if i < len(rows): # update xml row with values if the value exists in rows array
                xmlCol.text = str(rows[i][j])
                j += 1 # next xml row
            else: # if no more data in rows fill remaining xml rows with blank data
                xmlCol.text = ''          
        i += 1 # next row in xml doc
    tree.write(outputPath[:fnInsertPos] + str(x//numberOfRows + 1) + outputPath[fnInsertPos:]); # write data to new xml file

print("Done!")
taxDocument.close()




