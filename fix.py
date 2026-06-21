"""
fix.py - Fix XML obligatoriu: adauga <w:tblInd w:w="0"/> ca tabelul sa poata fi
selectat complet in Word.

Utilizare:
    python fix.py FISIER.docx
    python fix.py                 # implicit: telc_B2_pagini_10_13.docx
"""
import sys
import zipfile

src = sys.argv[1] if len(sys.argv) > 1 else 'telc_B2_pagini_10_13.docx'

with zipfile.ZipFile(src, 'r') as zin:
    files = {name: zin.read(name) for name in zin.namelist()}

xml = files['word/document.xml'].decode('utf-8')

if '<w:tblInd' in xml:
    print(f'{src}: tblInd deja prezent, nimic de facut.')
else:
    xml = xml.replace(
        '<w:tblW w:type="dxa" w:w="10800"/>',
        '<w:tblW w:type="dxa" w:w="10800"/><w:tblInd w:w="0" w:type="dxa"/>'
    )
    files['word/document.xml'] = xml.encode('utf-8')
    with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    print(f'Fix XML aplicat pe {src}.')
