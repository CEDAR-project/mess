#!/usr/bin/env python

# mess - Microsoft Excel Style-extractor and Styler

from xlutils.styles import Styles
from xlrd import open_workbook
import argparse
import os
import json

parser = argparse.ArgumentParser(description = "MESS - Microsoft Excel Style-extractor and Styler", usage = "mess.py ")
parser.add_argument('--input', '-i',
                    help = "Input XLS file to read styles from", 
                    required = True)
parser.add_argument('--output', '-o',
                    help = "Output XLS file to apply styles to", 
                    required = False)
parser.add_argument('--serialize', '-s',
                    help = "Output XLS file to apply styles to", 
                    required = False)

args = parser.parse_args()

# print args.input, args.serialize

# Styles data structure
# List of workbook sheets
# Each sheet is a dictionary, whose keys are coord tuples and values are styles
bookStyles = []

# Since json won't allow to serialize tuples as keys, we define this remap
def remap_keys(mapping):
    remap = []
    for sheet in mapping:
        remap.append([{'key':k, 'value': v} for k, v in sheet.iteritems()])
    return remap

book = open_workbook(args.input, formatting_info = 1)
styles = Styles(book)
for i in range(book.nsheets):
    sheetStyles = {}
    sheet = book.sheet_by_index(i)
    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            sheetStyles[(i,j)] = styles[sheet.cell(i,j)].name
    bookStyles.append(sheetStyles)

with open(args.serialize, 'w') as outfile:
    json.dump(remap_keys(bookStyles), outfile)
