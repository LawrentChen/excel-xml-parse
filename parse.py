#!usr/bin/env/python3
# -*- coding: utf-8 -*-

#A snippet to parse office xml format spreadsheet

# Enlightened by Jonathan Whitemore and jrovegno

import pandas as pd
import xml.sax
import os

# Get the list of files in working directory
workdir = '' # Input your directory here
filenames = os.listdir(workdir)
filepath = []
for i in range(0,len(filenames)):
    filepath.append(os.path.join(workdir,filenames[i]))

# Prepare the ContentHandler for parser
class ExcelHandler(xml.sax.ContentHandler):
    def __init__(self):
        self.chars = []
        self.cells = []
        self.rows = []
        self.tables = []
    def startElement(self, name, atts):
        if name=="Data":
            self.chars = []
        elif name=="Row":
            self.cells=[]
        elif name=="Table":
            self.rows = []
    def characters(self, content):
        self.chars.append(content)
    def endElement(self, name):
        if name=="Data":
            self.cells.append(''.join(self.chars))
        elif name=="Row":
            self.rows.append(self.cells)
        elif name=="Table":
            self.tables.append(self.rows)

# Iterate to get the content of all files
content = []

for file in filepath:
    excelHandler = ExcelHandler()
    xml.sax.parse(file, excelHandler)
    fileparse = pd.DataFrame(excelHandler.tables[0][3:-2], columns=excelHandler.tables[0][2]) # Change this slicer according to the spreadsheet structure if necessary

    # Append to list
    content.append(fileparse)

# Concatenate the files into DataFrame
result = pd.concat(content)
