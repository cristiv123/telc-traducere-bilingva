#!/usr/bin/env python3
"""
lectii.py - Helper determinist pentru manuale organizate pe "Lektion N"
(ex. Sicher C1.1: Arbeitsbuch + cheia AB Loesungen, doua PDF-uri separate).

Detecteaza numarul lectiei de pe fiecare pagina din titlul "Lektion N" / "LEKTION N".
Numarul lectiei trece corect prin OCR chiar cand restul titlului e corupt (verificat in
Faza 1). Paginile fara titlu detectabil mostenesc lectia ultimei pagini cu titlu
(carry-forward) - lectiile sunt contigue intr-un manual.

Subcomenzi:
    python lectii.py map    "PDF" START-END
        Listeaza per pagina lectia detectata:
            NN  L<k>          titlu gasit pe pagina
            NN  L<k>(cont)    fara titlu, mostenita prin carry-forward
            NN  ?             nedeterminabila (nicio lectie vazuta inca in interval)

    python lectii.py verify "PDF" START-END L3[,L4]
        Garda: exit 0 daca setul de lectii detectat (cu carry-forward) coincide cu lista
        data; altfel mesaj clar pe stderr + exit !=0. Paginile '?' fac verify sa esueze
        (cerinta: daca nu pot determina lectia unei pagini, opreste-te si cere explicit).

    python lectii.py find   "SOL_PDF" N
        Tipareste intervalul de pagini din cheie care contine solutiile pentru Lektion N:
        de la prima pagina cu "Lektion N" pana la prima pagina cu "Lektion N+1" INCLUSIV
        (ca sa prinda coada lectiei care curge pe pagina urmatoare). Ultima lectie -> final.
        Iesire: "START-END" (sau "START" daca o singura pagina). Exit !=0 daca nu o gaseste.

    python lectii.py plan   "SOL_PDF" L3[,L4]
        Tipareste pe o singura linie maparea lectie->pagini de cheie, gata de pus in prompt:
            "Lektion 3 -> cheie pag 4-7; Lektion 4 -> cheie pag 7-10"
        Exit !=0 daca vreo lectie nu e gasita in cheie. (Garda de potrivire cu intervalul
        de exercitii se face separat cu 'verify'.)

Numerele sunt pagini PDF (1-based).
"""

import sys
import re
import os

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("EROARE: PyMuPDF nu este instalat. Ruleaza: pip install pymupdf")

# "Lektion 3", "LEKTION  3", "Lektion03" -> 3. Numarul e robust la OCR; titlul nu conteaza.
LEKTION_RE = re.compile(r"(?i)\blektion\s*0*(\d+)\b")


def _open(pdf_path):
    if not os.path.isfile(pdf_path):
        sys.exit(f"EROARE: PDF-ul nu exista: {pdf_path}")
    return fitz.open(pdf_path)


def parseaza_interval(arg, total):
    arg = arg.strip()
    try:
        if "-" in arg:
            a, b = arg.split("-", 1)
            start, end = int(a), int(b)
        else:
            start = end = int(arg)
    except ValueError:
        sys.exit(f"EROARE: interval invalid '{arg}'. Exemplu: 35-50 sau 35")
    if start < 1 or end < start:
        sys.exit(f"EROARE: interval invalid '{arg}'.")
    if end > total:
        sys.exit(f"EROARE: pagina {end} depaseste totalul PDF ({total}).")
    return start, end


def lectie_pe_pagina(doc, n):
    """Numarul lectiei al carei titlu apare pe pagina PDF n (1-based), sau None."""
    m = LEKTION_RE.search(doc[n - 1].get_text())
    return int(m.group(1)) if m else None


def construieste_map(doc, start, end):
    """
    Returneaza lista de tuple (pagina, lectie_sau_None, este_cont).
    Carry-forward: pagina fara titlu mosteneste ultima lectie vazuta. Inainte de prima
    lectie din interval, lectia e None (nedeterminabila).
    """
    rezultat = []
    curenta = None
    for n in range(start, end + 1):
        det = lectie_pe_pagina(doc, n)
        if det is not None:
            curenta = det
            rezultat.append((n, det, False))
        elif curenta is not None:
            rezultat.append((n, curenta, True))
        else:
            rezultat.append((n, None, False))
    return rezultat


def normalizeaza_lectii(arg):
    """'L3', '3', 'L3,L4', 'l3, 4' -> set({3,4})."""
    out = set()
    for token in arg.split(","):
        token = token.strip().lower().lstrip("l")
        if not token:
            continue
        if not token.isdigit():
            sys.exit(f"EROARE: lista de lectii invalida: '{arg}'. Ex: L3 sau L3,L4")
        out.add(int(token))
    if not out:
        sys.exit(f"EROARE: lista de lectii goala: '{arg}'.")
    return out


def cmd_map(pdf, interval):
    doc = _open(pdf)
    start, end = parseaza_interval(interval, doc.page_count)
    for n, lec, cont in construieste_map(doc, start, end):
        if lec is None:
            print(f"{n}\t?")
        elif cont:
            print(f"{n}\tL{lec}(cont)")
        else:
            print(f"{n}\tL{lec}")
    doc.close()


def cmd_verify(pdf, interval, lista):
    doc = _open(pdf)
    start, end = parseaza_interval(interval, doc.page_count)
    asteptat = normalizeaza_lectii(lista)
    harta = construieste_map(doc, start, end)
    doc.close()

    nedeterminate = [n for n, lec, _ in harta if lec is None]
    detectate = {lec for _, lec, _ in harta if lec is not None}

    if nedeterminate:
        pag = ", ".join(str(n) for n in nedeterminate)
        sys.stderr.write(
            f"EROARE: nu pot determina lectia pentru pagina/paginile: {pag}.\n"
            f"Spune-mi explicit din ce Lektion fac parte (apar inainte de orice titlu "
            f"'Lektion N' din interval).\n")
        sys.exit(2)

    if detectate != asteptat:
        lipsa = asteptat - detectate
        in_plus = detectate - asteptat
        msg = "EROARE: lectiile detectate nu corespund listei --lectii.\n"
        msg += f"  detectate in pagini {start}-{end}: {sorted('L%d'%x for x in detectate)}\n"
        msg += f"  cerute (--lectii):            {sorted('L%d'%x for x in asteptat)}\n"
        if in_plus:
            msg += f"  in plus fata de lista: {sorted('L%d'%x for x in in_plus)} "
            msg += "(intervalul de pagini acopera lectii pe care nu le-ai trecut)\n"
        if lipsa:
            msg += f"  lipsa din pagini: {sorted('L%d'%x for x in lipsa)} "
            msg += "(le-ai cerut dar nu apar in interval)\n"
        sys.stderr.write(msg)
        sys.exit(3)

    print(f"OK: pagini {start}-{end} acopera exact {sorted('L%d'%x for x in detectate)}.")
    sys.exit(0)


def cmd_find(pdf, n_str):
    doc = _open(pdf)
    try:
        n = int(str(n_str).strip().lower().lstrip("l"))
    except ValueError:
        sys.exit(f"EROARE: numar de lectie invalid: '{n_str}'")
    total = doc.page_count

    start_pg = None
    next_pg = None
    for p in range(1, total + 1):
        det = lectie_pe_pagina(doc, p)
        if det == n and start_pg is None:
            start_pg = p
        if start_pg is not None and det is not None and det > n:
            next_pg = p
            break
    doc.close()

    if start_pg is None:
        sys.stderr.write(f"EROARE: Lektion {n} nu a fost gasita in {pdf}.\n")
        sys.exit(4)

    # Pana la pagina pe care apare lectia urmatoare INCLUSIV (coada lectiei curge pe ea),
    # sau pana la final daca e ultima lectie.
    end_pg = next_pg if next_pg is not None else total
    print(f"{start_pg}" if start_pg == end_pg else f"{start_pg}-{end_pg}")


def _span_lectie(doc, n, total):
    """Intervalul de pagini din cheie pentru Lektion n (vezi cmd_find). None daca lipseste."""
    start_pg = None
    next_pg = None
    for p in range(1, total + 1):
        det = lectie_pe_pagina(doc, p)
        if det == n and start_pg is None:
            start_pg = p
        if start_pg is not None and det is not None and det > n:
            next_pg = p
            break
    if start_pg is None:
        return None
    end_pg = next_pg if next_pg is not None else total
    return (start_pg, end_pg)


def cmd_plan(sol_pdf, lista):
    doc = _open(sol_pdf)
    total = doc.page_count
    lectii = sorted(normalizeaza_lectii(lista))
    bucati = []
    lipsa = []
    for n in lectii:
        sp = _span_lectie(doc, n, total)
        if sp is None:
            lipsa.append(n)
            continue
        a, b = sp
        span = f"{a}" if a == b else f"{a}-{b}"
        # separator ':' (nu '->'): '>' ar fi interpretat ca redirectare cand .bat-ul
        # scrie maparea intr-un fisier-prompt cu 'echo ... >> fisier'.
        bucati.append(f"Lektion {n}: cheie pag {span}")
    doc.close()
    if lipsa:
        pag = ", ".join(f"L{x}" for x in lipsa)
        sys.stderr.write(f"EROARE: lectiile {pag} nu au fost gasite in cheie {sol_pdf}.\n")
        sys.exit(4)
    print("; ".join(bucati))


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    cmd = sys.argv[1].lower()
    try:
        if cmd == "map":
            cmd_map(sys.argv[2], sys.argv[3])
        elif cmd == "verify":
            cmd_verify(sys.argv[2], sys.argv[3], sys.argv[4])
        elif cmd == "find":
            cmd_find(sys.argv[2], sys.argv[3])
        elif cmd == "plan":
            cmd_plan(sys.argv[2], sys.argv[3])
        else:
            sys.exit(f"EROARE: subcomanda necunoscuta '{cmd}'. Foloseste: map | verify | find")
    except IndexError:
        sys.exit("EROARE: argumente lipsa.\n" + __doc__)


if __name__ == "__main__":
    main()
