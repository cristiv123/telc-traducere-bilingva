"""
valid.py - Validare rapida a structurii XML a documentului generat.

Utilizare:
    python valid.py FISIER.docx
    python valid.py              # implicit: telc_B2_pagini_10_13.docx

Valori corecte: pgSz 12240x15840, tblW 10800, tblInd 0, tblLayout fixed,
doua gridCol 5400.
"""
import sys
import zipfile
import re

src = sys.argv[1] if len(sys.argv) > 1 else 'telc_B2_pagini_10_13.docx'

with zipfile.ZipFile(src) as z:
    xml = z.read('word/document.xml').decode('utf-8')
    print('fisier:', src)
    print('pgSz:', re.findall(r'<w:pgSz[^/]*/>', xml))
    print('tblW:', re.findall(r'<w:tblW[^/]*/>', xml))
    print('tblInd:', re.findall(r'<w:tblInd[^/]*/>', xml))
    print('tblLayout:', re.findall(r'<w:tblLayout[^/]*/>', xml))
    print('gridCols:', re.findall(r'<w:gridCol[^/]*/>', xml)[:3])
