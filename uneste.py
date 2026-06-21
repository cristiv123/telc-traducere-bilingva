"""
uneste.py - Uneste bucatile .docx dintr-un interval intr-un singur document,
in ordinea paginilor, pastrand tabelele bilingve (docxcompose).

Utilizare:
    python uneste.py "<base_fara_extensie>" START END
    ex: python uneste.py "Z:\\manual" 10 50

Reconstruieste numele bucatilor (cate 2 pagini; ultima impara = o pagina):
    <base>_pagini_LO_HI.docx   (interval de 2)
    <base>_pagini_N.docx       (o pagina)
Rezultat:
    <base>_COMPLET_START_END.docx
"""
import sys
import os

try:
    from docx import Document
    from docxcompose.composer import Composer
except ImportError:
    sys.exit("EROARE: lipseste docxcompose/python-docx. Ruleaza: pip install docxcompose")


def chunk_files(base, start, end):
    """Reconstruieste, in ordine, caile bucatilor pentru intervalul [start, end]."""
    files = []
    i = start
    while i <= end:
        hi = i + 1
        if hi > end:
            hi = end
        label = f"{i}" if i == hi else f"{i}_{hi}"
        files.append(f"{base}_pagini_{label}.docx")
        i += 2
    return files


def main():
    if len(sys.argv) != 4:
        sys.exit("Utilizare: python uneste.py \"<base_fara_extensie>\" START END")

    base = sys.argv[1]
    try:
        start = int(sys.argv[2])
        end = int(sys.argv[3])
    except ValueError:
        sys.exit("EROARE: START si END trebuie sa fie numere.")

    files = chunk_files(base, start, end)

    missing = [f for f in files if not os.path.isfile(f)]
    if missing:
        sys.exit("EROARE: lipsesc bucati, nu unesc:\n  " + "\n  ".join(missing))

    out = f"{base}_COMPLET_{start}_{end}.docx"

    master = Document(files[0])
    composer = Composer(master)
    for f in files[1:]:
        composer.append(Document(f))
    composer.save(out)

    print(out)


if __name__ == "__main__":
    main()
