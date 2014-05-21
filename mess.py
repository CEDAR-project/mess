#!/usr/bin/env python

# mess - Microsoft Excel Style-extractor and Styler

from xlutils.styles import Styles
from xlrd import open_workbook
import argparse
import os
import json

parser = argparse.ArgumentParser(description = "MESS - Microsoft Excel Style-extractor and Styler")
parser.add_argument('--input', '-i',
                    help = "Input file(s)", 
                    nargs = '+',
                    required = True)

args = parser.parse_args()

# print args.input

# Styles data structure
# List of workbook sheets
# Each sheet is a dictionary, whose keys are coord tuples and values are styles
bookStyles = []

for f in args.input:
    book = open_workbook(f, formatting_info = 1)
    styles = Styles(book)
    for i in range(book.nsheets):
        sheetStyles = {}
        sheet = book.sheet_by_index(i)
        for i in range(sheet.nrows):
            for j in range(sheet.ncols):
                sheetStyles[(i,j)] = styles[sheet.cell(i,j)].name
        bookStyles.append(sheetStyles)

print(json.dumps(bookStyles, indent=4))
        

