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
    rows.append(row)

numberOfRows = 14 # set to how many rows in the table in the destination document


for x in range(0, len(rows), numberOfRows):
    i = x        
    for xmlRow in root.find('Page1').find('Table_Line1'):
        
        j = 0
        for xmlCol in xmlRow:
            xmlCol.text = str(rows[i][j])
            j += 1
            
        i += 1
    tree.write(outputPath[:fnInsertPos] + str(x//numberOfRows + 1) + outputPath[fnInsertPos:]);

print("Done!")
taxDocument.close()




