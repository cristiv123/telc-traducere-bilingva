---
name: telc-bilingual-docx
description: "Creează documente Word (.docx) bilingve germană-română din conținut extras din manuale de pregătire pentru examenul telc Deutsch B2 (sau alte manuale de germană). Folosește acest skill ori de câte ori utilizatorul încarcă pagini dintr-un manual de germană (telc, Kontext, Weitblick, Ach so etc.) și cere traducerea sau digitizarea lor într-un document bilingv tabelar. Triggere: 'fă același lucru cu paginile X până la Y', 'traduce paginile', 'document bilingv', 'tabel DE/RO', 'digitizează manualul', 'traducere din manual'. Skill-ul produce tabele Word cu două coloane egale (germană stânga, română dreapta), cu rânduri de secțiune gri, bold pe termeni specializați, toate caracterele germane corecte (ä ö ü ß) și fix-urile XML necesare pentru selecția corectă a tabelului în Word. Se aplică și pentru transcrieri de examene orale, texte de citire, exerciții de gramatică și orice alt conținut dintr-un manual de limbă germană."
---

# Skill: Document bilingv DE/RO din manual telc B2

## Scop

Generează un document Word (.docx) cu tabel bilingv german–român, propoziție cu propoziție, din orice pagini ale unui manual de germană furnizate de utilizator (ca PDF, imagini sau text copiat).

## Mediu (Windows / Claude Code)

Acest skill rulează pe Windows prin Claude Code. Note importante:
- Pachetul `docx` npm trebuie să fie instalat global: verifică cu `npm list -g docx`. Dacă lipsește: `npm install -g docx`.
- Fișierele de output se scriu în **folderul de lucru curent** (de ex. `.\telc_B2_pagini_XX_YY.docx`), NU în `/mnt/user-data/outputs/`.
- Pentru a citi un PDF e nevoie de poppler (`pdftotext`, `pdftoppm`) sau de o bibliotecă Python (`pdfplumber` / `pypdf`). Dacă poppler nu e instalat pe Windows, extrage textul cu Python: `pip install pdfplumber` și citește paginile cerute.
- La rularea unui script Node global instalat, dacă apare eroare „Cannot find module docx", setează `NODE_PATH` către folderul global npm: pe Windows `set NODE_PATH=%APPDATA%\npm\node_modules` (sau rezultatul comenzii `npm root -g`).

## Workflow obligatoriu

> ⚠️ **AVERTISMENT — numerotarea paginilor.** Numărul paginii din fișierul PDF de regulă NU coincide cu cel tipărit în manual: **pagina N din PDF = pagina tipărită N−1** (copertă/pagini de gardă nenumerotate). Exemplu real: pagina PDF 18 = pagina tipărită 17. **Înainte de a extrage sau traduce, cere mereu confirmarea**: utilizatorul se referă la numărul din PDF sau la cel tipărit pe pagină? Nu presupune — diferența de 1 duce la traducerea paginii greșite.

### Pasul 1 — Citește conținutul sursă

Dacă utilizatorul a furnizat un PDF sau imagini: extrage textul complet, în ordinea apariției. Nu omite niciun element textual: instrucțiuni, sfaturi metodice, sarcini, texte de citire/ascultare, întrebări, răspunsuri, explicații ale soluțiilor, exemple de scriere.

Dacă textul vine din OCR (cum e cazul manualului scanat), **compararea textului OCR cu imaginea paginii este OBLIGATORIE, nu opțională** — OCR-ul greșește sistematic: coloane citite intercalat, cifre citite ca litere („1." → „in", „a-l" → „a-I"), cuvinte trunchiate, diacritice pierdute.

Procedură:
1. Extrage **ambele** (text + PNG la 150 dpi) cu scriptul `extrage.py` (PyMuPDF), care le pune în subfolderul `pagini_extrase/` ca `pagina_NN.txt` + `pagina_NN.png`. (`pdftoppm`/poppler NU e disponibil pe acest Windows; PyMuPDF/`fitz` este deja instalat și randează singur PNG-ul.)
2. **Citește efectiv** imaginea `pagini_extrase/pagina_NN.png` cu tool-ul Read și confruntă fiecare pasaj cu textul `.txt` înainte de traducere.
3. Corectează: ordinea de citire a textelor pe 2 coloane, numerotarea pierdută, cifrele/literele confundate, cuvintele trunchiate; redă fidel chiar și greșelile tipărite ale manualului (semnalează-le utilizatorului).

### Pasul 2 — Identifică structura

Grupează conținutul în secțiuni logice:
- Marcatori de pagină și subsecțiune (ex. „PAGINA 25 — Sprachbausteine Teil 1")
- Instrucțiuni de timp și metodă
- Instrucțiunea sarcinii
- Textul propriu-zis (scrisori, articole, anunțuri, transcrieri)
- Variante de răspuns
- Sfaturi metodice (căsuțe cu beculeț)
- Explicații ale soluțiilor (Die Lösungen verstehen)

### Pasul 3 — Construiește fișierul JS și generează .docx

Folosește **docx npm** (disponibil global: `npm list -g docx`). Scrie un script Node.js care:

> 🧹 **Unde se scrie scriptul generator.** Scriptul de generare este **temporar** și se scrie
> **întotdeauna** în subfolderul `_temp\` al folderului de lucru (creează-l dacă lipsește), cu nume
> de forma `_temp\gen_pagini_XX_YY.js`. NU scrie scripturi de generare la nivelul folderului de
> lucru (acolo stau doar uneltele permanente: `extrage.py`, `fix.py`, `valid.py`, `uneste.py`,
> `telc.bat`, `telc-batch.bat`). `_temp\` e scratch de unică folosință — se poate goli oricând fără
> să atingi uneltele. **`.docx`-ul final NU se scrie niciodată în `_temp\`** — vezi „Denumirea
> fișierelor output".

1. Importă: `Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, BorderStyle, WidthType, ShadingType, TableLayoutType`

2. Definește funcțiile helper standard:

```javascript
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        BorderStyle, WidthType, ShadingType, TableLayoutType } = require('docx');

const border = { style: BorderStyle.SINGLE, size: 8, color: '000000' };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function parseRuns(text) {
  const parts = text.split(/(\*\*[^*]+\*\*)/g);
  return parts.map(p => {
    if (p.startsWith('**') && p.endsWith('**')) {
      return new TextRun({ text: p.slice(2,-2), bold: true, font: 'Arial', size: 20 });
    }
    return new TextRun({ text: p, font: 'Arial', size: 20 });
  });
}

function contentRow(de, ro) {
  return new TableRow({ children: [
    new TableCell({ borders, width: { size: 5400, type: WidthType.DXA }, margins: cellMargins,
      children: [new Paragraph({ children: parseRuns(de) })] }),
    new TableCell({ borders, width: { size: 5400, type: WidthType.DXA }, margins: cellMargins,
      children: [new Paragraph({ children: parseRuns(ro) })] }),
  ]});
}

function sectionRow(label) {
  return new TableRow({ children: [
    new TableCell({ borders, columnSpan: 2, width: { size: 10800, type: WidthType.DXA },
      margins: cellMargins, shading: { fill: 'D0D0D0', type: ShadingType.CLEAR },
      children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, font: 'Arial', size: 20 })] })] })
  ]});
}

function headerRow() {
  return new TableRow({ children: [
    new TableCell({ borders, width: { size: 5400, type: WidthType.DXA }, margins: cellMargins,
      shading: { fill: 'A0A0A0', type: ShadingType.CLEAR },
      children: [new Paragraph({ children: [new TextRun({ text: 'Germana (original)', bold: true, font: 'Arial', size: 22 })] })] }),
    new TableCell({ borders, width: { size: 5400, type: WidthType.DXA }, margins: cellMargins,
      shading: { fill: 'A0A0A0', type: ShadingType.CLEAR },
      children: [new Paragraph({ children: [new TextRun({ text: 'Romana (traducere)', bold: true, font: 'Arial', size: 22 })] })] }),
  ]});
}

// Randul de titlu al unui Teil/Aufgabe: o celula pe ambele coloane, gri inchis,
// bold, mai mare. Marcheaza inceputul unui TABEL NOU (vezi Pasul 4).
function teilTitleRow(label) {
  return new TableRow({ children: [
    new TableCell({ borders, columnSpan: 2, width: { size: 10800, type: WidthType.DXA },
      margins: cellMargins, shading: { fill: '808080', type: ShadingType.CLEAR },
      children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, font: 'Arial', size: 24 })] })] })
  ]});
}
```

3. Construiește array-ul `DATA` cu intrări de tipul:
   - `['teil', 'HÖREN Teil 1']` — **graniță de tabel nou**: începe un TABEL SEPARAT pentru
     fiecare Teil / Aufgabe / probă principală de examen. Eticheta devine randul de titlu
     (gri închis) al noului tabel. Vezi mai jos „Un tabel separat per Teil/Aufgabe".
   - `['section', 'Eticheta sectiunii']` — rând gri de subsecțiune **în interiorul** tabelului
     curent (instrucțiune, „Afirmatiile 41-45", „TIPPS", „Info" etc.). NU începe un tabel nou.
   - `['row', 'Text german', 'Traducere română']` — rând de conținut în tabelul curent.

4. Construiește documentul — **un tabel separat per Teil/Aufgabe** (vezi regula de mai jos).
   Fiecare `['teil', ...]` închide tabelul curent, adaugă un **paragraf gol** (ca tabelele
   să nu fuzioneze în Word) și deschide un tabel nou cu randul de titlu + antet:

```javascript
const children = [];
let currentRows = null;

function flushTable() {
  if (currentRows && currentRows.length) {
    children.push(new Table({
      width: { size: 10800, type: WidthType.DXA },
      columnWidths: [5400, 5400],
      layout: TableLayoutType.FIXED,
      rows: currentRows,
    }));
    // Paragraf gol intre tabele: in Word doua tabele lipite fuzioneaza intr-unul singur.
    children.push(new Paragraph({ children: [] }));
  }
  currentRows = null;
}

function startTable(title) {
  flushTable();
  currentRows = [];
  if (title) currentRows.push(teilTitleRow(title));
  currentRows.push(headerRow());
}

for (const entry of DATA) {
  if (entry[0] === 'teil') startTable(entry[1]);
  else if (entry[0] === 'section') { if (!currentRows) startTable(null); currentRows.push(sectionRow(entry[1])); }
  else { if (!currentRows) startTable(null); currentRows.push(contentRow(entry[1], entry[2])); }
}
flushTable();

const doc = new Document({
  styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 720, right: 720, bottom: 720, left: 720 }
      }
    },
    children: children
  }]
});

// Output în folderul de lucru curent (Windows). Schimbă numele după nevoie.
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('telc_B2_pagini_XX_YY.docx', buf);
  console.log('Done');
});
```

5. Rulează scriptul din `_temp\`, cu `NODE_PATH` setat: `NODE_PATH="$(npm root -g)" node _temp/gen_pagini_XX_YY.js`. Dacă apare „Cannot find module docx", `docx` e instalat global — setează `NODE_PATH` către rezultatul `npm root -g` (pe Windows cmd: `set NODE_PATH=%APPDATA%\npm\node_modules`), apoi rulează din nou `node _temp\gen_pagini_XX_YY.js`.

#### Un tabel separat per Teil/Aufgabe (REGULĂ DE STRUCTURĂ)

**Fiecare Teil / Aufgabe / probă principală de examen trebuie să fie un TABEL SEPARAT**, nu
rânduri într-un singur tabel uriaș. Exemple de granițe care cer un `['teil', ...]` nou:
`HÖREN Teil 1`, `HÖREN Teil 2`, `LESEN Teil 1`, `Sprachbausteine Teil 1`,
`Schriftlicher Ausdruck`, `Mündlicher Ausdruck Teil 1` etc.

- Pune un `['teil', 'ETICHETA']` **la începutul fiecărui Teil/Aufgabe**. Eticheta devine
  randul de titlu (gri închis `808080`, bold, size 24) al noului tabel.
- Subsecțiunile din interiorul aceluiași Teil (instrucțiune, „Info", „TIPPS",
  „Afirmatiile 41-45", explicații) rămân `['section', ...]` — rânduri gri **în același tabel**.
- Între tabele scriptul inserează automat un **paragraf gol**; fără el, Word fuzionează două
  tabele alăturate într-unul singur.
- **Un Teil care se întinde pe mai multe bucăți de 2 pagini** poate apărea ca două tabele
  (câte unul per bucată) — acest lucru e **acceptat**; nu încerca să le unești peste bucăți.
  Important e doar ca **în interiorul unei bucăți** fiecare Teil să fie tabel separat.

### Pasul 4 — Fix XML obligatoriu (selecție tabel în Word)

**Întotdeauna** aplică acest fix după generare, altfel utilizatorul nu poate selecta întreg tabelul în Word. Rulează cu `python fix.py` (ajustează numele fișierului):

```python
import zipfile

src = 'telc_B2_pagini_XX_YY.docx'

with zipfile.ZipFile(src, 'r') as zin:
    files = {name: zin.read(name) for name in zin.namelist()}

xml = files['word/document.xml'].decode('utf-8')
xml = xml.replace(
    '<w:tblW w:type="dxa" w:w="10800"/>',
    '<w:tblW w:type="dxa" w:w="10800"/><w:tblInd w:w="0" w:type="dxa"/>'
)

files['word/document.xml'] = xml.encode('utf-8')

with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in files.items():
        zout.writestr(name, data)

print('Fix XML aplicat.')
```

**De ce este necesar:** `tblLayout type="fixed"` (setat de `TableLayoutType.FIXED`) împiedică redimensionarea automată a coloanelor. `tblInd w="0"` setează explicit indentarea la zero. Fără ambele, Word tratează tabelul ca parțial în afara zonei de text și nu permite selecția completă.

### Pasul 5 — Prezintă fișierul

Pe Windows/Claude Code: anunță utilizatorul calea completă a fișierului generat (de ex. `C:\claude-lab\02-traducere-telc\telc_B2_pagini_XX_YY.docx`) ca să-l poată deschide în Word.

---

## Reguli de conținut

### Traducere
- **Propoziție cu propoziție** — fiecare propoziție pe rând separat în `DATA`
- **Traducere integrală** — nu se omite niciun element textual
- Limbaj natural în română, fără calchieri forțate
- Termenii tehnici de examen păstrați consecvent:
  - Leseverstehen → Citire
  - Hörverstehen → Ascultare
  - Sprachbausteine → Elemente lingvistice
  - Schriftlicher Ausdruck → Exprimare scrisă
  - Mündlicher Ausdruck → Exprimare orală
  - Antwortbogen → foaie de răspuns
  - Aufgabe → sarcină
  - Lösungen → soluții
  - Überschrift → titlu
  - Teilnehmer/in → candidat/candidată
  - Prüfer/in → examinator/examinatoare
  - Schlüsselwörter → cuvinte-cheie
  - falsche Fährten → piste false
  - Situationsbeschreibungen → descrieri de situatii

### Diacritice românești
Conform preferinței utilizatorului, în coloana română **NU se folosesc diacritice** (ă â î ș ț). Scrie „solutie" nu „soluție", „romana" nu „română", „sarcina" nu „sarcină". Caracterele germane (ä ö ü ß) rămân întotdeauna corecte în coloana germană.

### Bold (`**cuvânt**`)
Marchează cu bold în **ambele coloane** termenii specializați sau compusele dificile:
- Termeni psihologici/medicali: Burnout, Erschöpfung / epuizare, Hypothese / ipoteza
- Termeni economici: Work-Life-Balance, Kundenbindung / fidelizare a clientilor
- Termeni de cercetare: Langzeitstudie / studiu longitudinal, Probanden / subiecti
- Compuse germane dificile: Leistungsfähigkeit, Vertrauenswürdigkeit, Ausbildungsleitlinien
- Expresii fixe: aus dem Bauch heraus / din instinct, in Anspruch nehmen / a utiliza
- Roluri în dialog: **Prüferin:** / **Examinatoare:**, **TN A:** / **Candidata A:**

### Caractere germane
Folosește întotdeauna caracterele corecte direct în string-urile JS: ä ö ü Ä Ö Ü ß
Nu folosi secvențe de escape `\uXXXX` în fișierele JS (pot cauza erori de sintaxă dacă nu sunt procesate corect).

### Ghilimele germane
Păstrează în coloana germană: „..." și '...' (ghilimele specifice limbii germane). Dacă ghilimelele curbe rup string-urile JS, folosește backtick (`) pentru string sau escapează corect.

### Structura rândurilor de secțiune
Etichetele de secțiune includ:
- Numărul/intervalul de pagini: `PAGINA 25 — Modelltest 1 · Sprachbausteine Teil 1`
- Subsecțiunile: `Instructiunea sarcinii`, `Afirmatiile 41-45`, `Sfat metodic general`
- Explicații: `Explicatia solutiei 21b — von meinen ersten Wochen berichten`

---

## Documente exercitiu + solutii (in acelasi tabel cu Teil-ul lor)

Cand utilizatorul cere si solutiile, pune **solutiile in ACELASI tabel cu exercitiul (Teil-ul)
de care apartin**, ca o continuare, nu intr-un tabel separat. (Regula „un tabel per Teil"
ramane valabila: solutiile lui Teil 1 stau in tabelul lui Teil 1, deci NU pui un `['teil', ...]`
nou pentru solutii — folosesti `['section', 'SOLUTII / LÖSUNGEN — …']`.) Solutiile
(Lösungsschlüssel) sunt de regula intr-o sectiune separata la finalul manualului
(ex. „LÖSUNGEN MODELLTEST 1", la pagina PDF 82), nu langa exercitiu.

Structura `DATA`:
1. Sectiunile exercitiului (ca de obicei): instructiune, titluri/variante, texte.
2. O sectiune finala de tip `['section', 'SOLUTII / LÖSUNGEN — <Prüfungsteil> · <Teil>']`.
3. Sub ea: nota introductiva a cheii (daca exista), **cheia compacta** (ex.
   `Lösungen: 1 **i**, 2 **a**, 3 **f**, 4 **e**, 5 **g**`), apoi explicatiile
   `Zu 1:` … `Zu N:` (germana + romana, propozitie cu propozitie).

Reguli specifice solutiilor:
- **Cheia compacta este aproape mereu corupta de OCR** (ex. „1 i" → „11",
  „2 a" → „2a" lipit). Verific-o OBLIGATORIU pe imaginea paginii de solutii si
  **coreleaz-o cu explicatiile** `Zu N:` si cu continutul exercitiului inainte de a o
  scrie. Litera corecta a solutiei reiese din explicatie (care numeste si „capcana").
- **Bold** pe litera-solutie in cheie si pe **litera-capcana** mentionata in fiecare
  explicatie (in ambele coloane), ex. `Auch bei **j** geht es um…` / `Si la **j**…`.
- Extrage DOAR partea ceruta din pagina de solutii (ex. „Prüfungsteil Leseverstehen /
  Teil 1"), nu toata pagina — o pagina de Lösungen contine de obicei solutiile mai
  multor parti (Teil 1, Teil 2, Fokus 1, Fokus 2 etc.).
- Eticheta `Zu N:` din germana → `La N:` in romana.

---

## Gestionarea erorilor frecvente

### Cannot find module 'docx'
Cauza: scriptul rulează cu Node local dar `docx` e instalat global.
Soluție: pe Windows `set NODE_PATH=%APPDATA%\npm\node_modules && node script.js`, sau folosește rezultatul `npm root -g`. Alternativ instalează `docx` local în folder: `npm install docx`.

### Tabelul nu poate fi selectat complet în Word
Cauza: lipsă `tblInd` și/sau `tblLayout` nu e `fixed`.
Soluție: aplică întotdeauna fix-ul XML din Pasul 4.

### PDF-ul nu se citește
Cauza: poppler (`pdftotext`) nu e instalat pe Windows.
Soluție: `pip install pdfplumber` și extrage textul paginilor cerute cu Python, sau instalează poppler pentru Windows.

---

## Denumirea fișierelor output

**`.docx`-ul final** se scrie în folderul de lucru curent (sau în `Z:\`, dacă așa cere fluxul),
NU în `_temp\`. **Scriptul generator** se scrie în `_temp\` (vezi Pasul 3). Astfel folderul de
lucru rămâne curat — doar uneltele permanente + documentele finale.

Convenție `.docx`: `telc_B2_pagini_XX_YY.docx` sau un nume descriptiv:
- `telc_B2_pagini_16_19.docx`
- `telc_B2_pagini_25_30.docx`
- `telc_B2_Transkriptionen_Mundliche_Prufung.docx`

**Curățenie scratch** (sigură oricând, nu atinge uneltele): șterge conținutul `_temp\`, ex.
`rm -rf _temp/*` (bash) sau `Remove-Item _temp\* -Force` (PowerShell).

---

## Validare rapidă

După generare, verifică structura XML cu `python valid.py`:
```python
import zipfile, re
with zipfile.ZipFile('telc_B2_pagini_XX_YY.docx') as z:
    xml = z.read('word/document.xml').decode('utf-8')
    print('pgSz:', re.findall(r'<w:pgSz[^/]*/>', xml))
    print('tblW:', re.findall(r'<w:tblW[^/]*/>', xml))
    print('tblInd:', re.findall(r'<w:tblInd[^/]*/>', xml))
    print('tblLayout:', re.findall(r'<w:tblLayout[^/]*/>', xml))
    print('gridCols:', re.findall(r'<w:gridCol[^/]*/>', xml)[:3])
```

Valorile corecte:
- `pgSz`: `w:w="12240" w:h="15840"` (US Letter)
- `tblW`: `w:w="10800"` (= 12240 - 720×2 margini)
- `tblInd`: `w:w="0"` ← **obligatoriu**
- `tblLayout`: `type="fixed"` ← **obligatoriu**
- `gridCols`: două intrări `w:w="5400"` (coloane egale)
