# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Ce este acest folder

Spațiu de lucru pentru digitizarea manualului telc Deutsch B2 (`Dagmar Giersberg - Prüfung Express – telc Deutsch B2 - 2021_ocr.pdf`, 98 pagini) în documente Word bilingve germană–română, tabelare. Întreg fluxul este condus de skill-ul `telc-bilingual-docx` (definit în `C:\Users\cristiv\.claude\skills\telc-bilingual-docx\SKILL.md`) — consultă-l înainte de a genera un document nou; conține șablonul JS canonic, regulile de traducere și dicționarul de termeni.

## Pipeline-ul de lucru (ordinea contează)

1. **Extragere** — `python extrage.py 14-19` (sau o singură pagină `python extrage.py 18`). Scrie în `pagini_extrase/` câte un `pagina_NN.txt` (text brut OCR) **și** un `pagina_NN.png` (imaginea la 150 dpi) pentru fiecare pagină. Folosește PyMuPDF (`fitz`).
2. **Corectare OCR — OBLIGATORIE** — citește imaginea `pagini_extrase/pagina_NN.png` cu tool-ul Read și confruntă fiecare pasaj cu `.txt`-ul. OCR-ul greșește sistematic: coloane citite intercalat, cifre ca litere („1."→„in", „a-l"→„a-I"), cuvinte trunchiate, diacritice pierdute.
3. **Generare docx** — un script Node **temporar** per livrare, scris în subfolderul `_temp\` (ex. `_temp\gen_pagini_20_25.js`), **niciodată** la nivelul folderului de lucru (acolo stau doar uneltele permanente + documentele finale). Conține un array `DATA` cu rânduri `['row', textDE, textRO]` și `['section', eticheta]`, plus helperii standard (`parseRuns`, `contentRow`, `sectionRow`, `headerRow`). `.docx`-ul final se scrie în folderul de lucru (sau `Z:\`), NU în `_temp\`. Rulează cu `NODE_PATH` setat (vezi mai jos). `_temp\` e scratch de unică folosință — se poate goli oricând fără să atingi uneltele.
4. **Fix XML — OBLIGATORIU** — `python fix.py` adaugă `<w:tblInd w:w="0" .../>`; fără el tabelul nu se poate selecta complet în Word.
5. **Validare** — `python valid.py` verifică structura XML.

## Documente exercitiu + solutii (un singur tabel)

Skill-ul poate genera acum un document care combină exercițiul **și** soluțiile lui în
**același tabel**, ca o continuare (nu două tabele). Soluțiile (Lösungsschlüssel) sunt
într-o secțiune separată la finalul manualului — de regulă „LÖSUNGEN MODELLTEST …" pe
paginile PDF ~82-91 — nu lângă exercițiu, deci trebuie extrase separat și **corelate cu
exercițiul corespunzător**.

Structura: secțiunile exercițiului, apoi o secțiune `SOLUTII / LÖSUNGEN — <Prüfungsteil> ·
<Teil>`, sub care vin cheia compactă + explicațiile `Zu N:` (→ `La N:` în RO), traduse.
Extrage **doar** partea cerută din pagina de soluții (o pagină Lösungen conține soluțiile
mai multor părți: Teil 1, Teil 2, Fokus 1 etc.).

- **Capcană — cheia compactă coruptă de OCR**: cheia de răspunsuri (ex. `1 i, 2 a, 3 f,
  4 e, 5 g`) e aproape mereu citită greșit (devine `11,2a,3f,4e,5g`). **Litera corectă se
  deduce din explicația `Zu N:`**, care numește și „capcana" (titlul asemănător dar
  greșit). Verifică mereu pe imaginea paginii de soluții și corelează cu exercițiul
  înainte de a scrie cheia. Bold pe litera-soluție și pe litera-capcană în ambele coloane.

## Comenzi

```bash
# Extragere pagini (PDF page numbers, vezi avertismentul de mai jos)
python extrage.py 14-19

# Generare docx — scriptul temporar sta in _temp\; docx npm e instalat GLOBAL, deci NODE_PATH e obligatoriu:
NODE_PATH="$(npm root -g)" node _temp/gen_pagini_XX_YY.js

# Fix XML + validare (accepta numele fisierului ca argument; implicit telc_B2_pagini_10_13.docx)
python fix.py FISIER.docx
python valid.py FISIER.docx

# Curatenie scratch (sigura oricand, nu atinge uneltele permanente):
rm -rf _temp/*
```

Valorile corecte la validare: `pgSz` 12240×15840, `tblW` 10800, `tblInd` 0, `tblLayout` fixed, două `gridCol` 5400.

## Automatizare end-to-end (un singur apel)

Pentru a rula tot pipeline-ul (extragere → corectare OCR → traducere → fix XML → salvare)
fără pași manuali, există două batch-uri care invocă Claude Code neinteractiv
(`claude --dangerously-skip-permissions -p ...`). PDF-ul sursă se află în `Z:\`, iar
rezultatele se scriu tot în `Z:\` (`<base>_pagini_<pagini>.docx`).

```bat
:: o singura livrare (o pagina sau un interval); deschide rezultatul in LibreOffice
telc.bat "telc_b2.pdf" 18
telc.bat "telc_b2.pdf" 14-19

:: procesare RELUABILA pe bucati de 2 pagini, cu unire finala intr-un singur docx;
:: la limita de tokeni se opreste curat, iar reluarea aceleiasi comenzi continua de unde a ramas
telc-batch.bat "telc_b2.pdf" 10-50
telc-batch.bat "telc_b2.pdf"          :: tot fisierul
```

- **Nume PDF fara diacritice / caractere speciale** — `cmd.exe` corupe diacriticele din
  numele fisierului. Redenumeste sursa ASCII (ex. `telc_b2.pdf`) inainte de a rula.
- `telc-batch.bat` tine evidenta progresului in `Z:\<base>_progres.txt`; o bucata e
  considerata gata doar daca e in progres **și** documentul `.docx` exista.
- **`uneste.py`** uneste bucatile unui interval intr-un singur document, in ordinea
  paginilor, cu `docxcompose`: `python uneste.py "Z:\<base>" START END` →
  `<base>_COMPLET_START_END.docx` (apelat automat de `telc-batch.bat` la final).
- **`extrage.py` accepta calea PDF-ului ca al doilea argument**
  (`python extrage.py 14-19 "Z:\telc_b2.pdf"`); fara el, auto-detectie in folderul curent.

## Capcane esențiale

- **Numerotare pagini: pagina N din PDF = pagina tipărită N−1** în manual (ex. pagina PDF 18 = pagina tipărită 17). `extrage.py` folosește numere PDF (1-based). Întreabă mereu utilizatorul la ce numerotare se referă.
- **`NODE_PATH` obligatoriu** la rularea scripturilor Node: `docx` e instalat global, nu local. Fără `NODE_PATH="$(npm root -g)"` apare „Cannot find module docx".
- **`fix.py` și `valid.py` acceptă numele fișierului ca argument** (`python fix.py telc_B2_pagina_18.docx`); fără argument folosesc implicit `telc_B2_pagini_10_13.docx`. `fix.py` e idempotent — sare peste fișierele care au deja `tblInd`.
- **`pdftoppm`/poppler NU e disponibil** pe acest Windows (doar `pdftotext` din `/mingw64/bin`). Pentru imagini folosește PyMuPDF (`fitz`), deja instalat. `pdftotext` necesită `-enc UTF-8` ca să nu corupă ä ö ü ß.
- **Encoding consolă**: numele PDF apare „stricat" în stdout Windows (cp1252), dar fișierele scrise sunt corecte — ignoră.

## Reguli de conținut (rezumat; detaliile în SKILL.md)

- **Coloana română NU folosește diacritice** (scrie „solutie", „romana", „sarcina"). Caracterele germane (ä ö ü Ä Ö Ü ß) rămân întotdeauna corecte în coloana germană.
- Traducere **propoziție cu propoziție**, integrală; fiecare propoziție pe rând separat în `DATA`.
- Bold cu `**...**` în ambele coloane pentru termeni specializați.
- Redă fidel inclusiv greșelile tipărite ale manualului (ex. „Partnerin"→„Partner", „zu verzichten"→„verzichten" pe paginile 10-13), dar semnalează-le utilizatorului.

## Output

- Documentele `.docx` se scriu în folderul curent.
- La cerere, se copiază și în `Z:\aaa` (uneori cu sufixe de versiune, ex. `_v1_necorectat` / `_v2_corectat`).
- Deschidere pentru inspecție: LibreOffice la `C:\Program Files\LibreOffice\program\soffice.exe`.
