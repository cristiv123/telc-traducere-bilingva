#!/usr/bin/env python3
"""
extrage.py - Extrage text brut + imagine PNG din PDF-ul telc, pe interval de pagini.

Utilizare:
    python extrage.py 14-19                # paginile PDF 14 pana la 19 (inclusiv)
    python extrage.py 14                   # o singura pagina
    python extrage.py 14-19 "Z:\\x.pdf"    # PDF explicit (altfel auto-detectie in folder)

Pentru fiecare pagina din interval scrie in subfolderul pagini_extrase/:
    pagina_NN.txt   - textul brut (UTF-8)
    pagina_NN.png   - imaginea paginii la 150 dpi

Numerele sunt pagini PDF (1-based). Foloseste PyMuPDF (fitz).
"""

import sys
import glob
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("EROARE: PyMuPDF nu este instalat. Ruleaza: pip install pymupdf")

DPI = 150
OUTDIR = Path("pagini_extrase")


def gaseste_pdf():
    """Gaseste PDF-ul din folderul curent (prefera numele cu 'telc')."""
    pdfs = sorted(glob.glob("*.pdf"))
    if not pdfs:
        sys.exit("EROARE: niciun fisier PDF in folderul curent.")
    for p in pdfs:
        if "telc" in p.lower():
            return p
    if len(pdfs) == 1:
        return pdfs[0]
    sys.exit("EROARE: mai multe PDF-uri gasite, niciunul cu 'telc' in nume:\n  "
             + "\n  ".join(pdfs))


def parseaza_interval(arg):
    """Transforma '14-19' sau '14' intr-o lista de numere (1-based)."""
    arg = arg.strip()
    try:
        if "-" in arg:
            start_s, end_s = arg.split("-", 1)
            start, end = int(start_s), int(end_s)
        else:
            start = end = int(arg)
    except ValueError:
        sys.exit(f"EROARE: interval invalid '{arg}'. Exemplu corect: 14-19 sau 14")
    if start < 1 or end < start:
        sys.exit(f"EROARE: interval invalid '{arg}'. Inceputul >= 1 si sfarsitul >= inceput.")
    return list(range(start, end + 1))


def main():
    if len(sys.argv) < 2:
        sys.exit("Utilizare: python extrage.py INTERVAL [CALE_PDF]\n"
                 "  ex: python extrage.py 14-19\n"
                 "  ex: python extrage.py 14-19 \"Z:\\manual.pdf\"")

    pagini = parseaza_interval(sys.argv[1])
    # argument 2 optional: calea PDF-ului. Fara el, auto-detectie in folderul curent.
    if len(sys.argv) > 2:
        pdf_path = sys.argv[2]
        import os
        if not os.path.isfile(pdf_path):
            sys.exit(f"EROARE: PDF-ul nu exista: {pdf_path}")
    else:
        pdf_path = gaseste_pdf()

    doc = fitz.open(pdf_path)
    total = doc.page_count
    print(f"PDF: {pdf_path}  ({total} pagini)")

    OUTDIR.mkdir(exist_ok=True)

    extrase = 0
    for n in pagini:
        if n > total:
            print(f"  ! pagina {n} depaseste totalul ({total}) - sarita")
            continue

        page = doc[n - 1]  # fitz e 0-based, argumentul e 1-based

        # (1) text brut
        text = page.get_text()
        txt_path = OUTDIR / f"pagina_{n:02d}.txt"
        txt_path.write_text(text, encoding="utf-8")

        # (2) imagine PNG la 150 dpi
        pix = page.get_pixmap(dpi=DPI)
        png_path = OUTDIR / f"pagina_{n:02d}.png"
        pix.save(png_path)

        print(f"  pagina {n:>3}: {txt_path.name} ({len(text)} caractere) + {png_path.name}")
        extrase += 1

    doc.close()
    print(f"Gata. {extrase} pagina/pagini extrase in {OUTDIR}/")


if __name__ == "__main__":
    main()
