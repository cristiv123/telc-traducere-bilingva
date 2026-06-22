# MANUAL SCRIPTURI — `telc.bat` & `telc-batch.bat`

Referință rapidă pentru generarea documentelor bilingve DE/RO din manualul telc.
Ambele scripturi cheamă Claude Code neinteractiv (`claude --dangerously-skip-permissions -p`)
și folosesc skill-ul `telc-bilingual-docx`.

---

## 0. Pe scurt — care script?

| Vreau să... | Folosesc | De ce |
|---|---|---|
| O livrare rapidă (o pagină sau un interval mic, dintr-un singur apel) | **`telc.bat`** | Un singur apel Claude pentru tot intervalul; simplu, rapid. |
| Un interval mare / tot fișierul, reluabil la limită de tokeni | **`telc-batch.bat`** | Sparge în bucăți de 2 pagini, ține evidența progresului, reia de unde a rămas și unește la final. |

**Diferența esențială:** `telc.bat` face **un singur** apel Claude pentru tot intervalul (dacă lovești
limita, pierzi livrarea). `telc-batch.bat` face **câte un apel per bucată de 2 pagini**, salvează fiecare
bucată și poate fi reluat fără să refacă ce e gata.

---

## 1. Cerințe importante (citește înainte de prima rulare)

- **Rulează dintr-un terminal NORMAL (cmd/PowerShell), NU din interiorul unei sesiuni Claude Code.**
  Limitare cunoscută: un batch care cheamă `claude` rulat *din* o sesiune Claude termină doar prima
  bucată; restul nu se procesează. Deschide un Command Prompt obișnuit și rulează acolo.
- **Numele PDF fără diacritice / caractere speciale.** `cmd.exe` corupe diacriticele din numele
  fișierului → Claude primește o cale greșită și apelul cade. Redenumește sursa ASCII, ex. `telc_b2.pdf`.
- **PDF-ul sursă trebuie să fie în `Z:\`.** Ambele scripturi caută `Z:\<nume.pdf>` și ies cu eroare
  dacă lipsește.
- Rezultatele se scriu tot în **`Z:\`**.

---

## 2. Capcana de numerotare a paginilor

> **Pagina N din PDF = pagina tipărită N−1** în manual (copertă + pagini de gardă nenumerotate).
> Exemplu: pagina **PDF 18** = pagina **tipărită 17**.

Ambele scripturi (și `extrage.py`) folosesc **numere de pagină din PDF (1-based)**.

**Cum aflu ce număr să dau:** dacă te uiți la numărul *tipărit* pe colțul paginii din manual și vrei
acea pagină, **adaugă 1** și dă rezultatul scriptului. (Pagină tipărită 27 → dai `28`.)
La dubiu, deschide PDF-ul și uită-te la indexul de pagină al vizualizatorului — acela e numărul corect.

---

## 3. `telc.bat` — o livrare rapidă

### Sintaxă
```bat
telc.bat "nume.pdf" <pagina | interval>
```
- `arg1` = numele PDF-ului din `Z:\` (obligatoriu, trebuie să se termine în `.pdf`)
- `arg2` = o pagină (`18`) sau un interval (`14-19`) — **obligatoriu**

### Exemple
```bat
telc.bat "telc_b2.pdf" 18        :: o singura pagina (PDF 18 = tiparit 17)
telc.bat "telc_b2.pdf" 14-19     :: un interval, intr-un singur apel
```
> Nu acceptă „tot fișierul" fără argument — pentru asta folosește `telc-batch.bat`.

### Ce face, pas cu pas
1. Validează argumentele și că `Z:\nume.pdf` există.
2. Derivă numele de ieșire: `Z:\<base>_pagini_<label>.docx` (vezi §5).
3. Cheamă Claude o singură dată: extrage paginile → corectează OCR pe imagini → traduce
   propoziție cu propoziție → aplică fix-ul XML → salvează la calea de ieșire.
4. Dacă fișierul a fost generat, îl **deschide automat în LibreOffice**; altfel iese cu eroare.

---

## 4. `telc-batch.bat` — intervale mari, reluabil, cu unire finală

### Sintaxă
```bat
telc-batch.bat "nume.pdf" N-M     :: un interval (ex. 10-50)
telc-batch.bat "nume.pdf" N       :: o singura pagina (ex. 18)
telc-batch.bat "nume.pdf"         :: TOT fisierul
```
- Fără `arg2` = tot fișierul: scriptul citește numărul de pagini din PDF și pornește de la pagina 1.
  Înainte de a porni cere o **confirmare** (`pause`), pentru că poate dura.

### Exemple
```bat
telc-batch.bat "telc_b2.pdf" 28-29
telc-batch.bat "telc_b2.pdf" 10-50
telc-batch.bat "telc_b2.pdf"
```

### Ce face, pas cu pas
1. Validează PDF-ul și intervalul; determină `START`–`END` și numărul de bucăți (`TOTAL`).
2. Parcurge intervalul în **bucăți de 2 pagini** (ultima poate fi de 1 pagină dacă intervalul e impar).
3. Pentru fiecare bucată:
   - Dacă e deja gata (vezi §6), **o sare**.
   - Altfel scrie promptul în `_temp\prompt_<label>.txt`, îl trimite lui Claude prin STDIN și
     salvează bucata la `Z:\<base>_pagini_<label>.docx`.
4. La final, dacă toate bucățile sunt OK, **unește** documentele cu `uneste.py` în
   `Z:\<base>_COMPLET_START_END.docx`, aplică `fix.py` pe rezultat și îl **deschide în LibreOffice**.
   (Dacă a fost o singură bucată, aceea *este* documentul final, fără unire.)

---

## 5. Unde se salvează și cum se numesc fișierele (toate în `Z:\`)

Fie `<base>` = numele PDF-ului fără `.pdf` (ex. `telc_b2.pdf` → `telc_b2`).

| Fișier | Cine îl creează | Ce e |
|---|---|---|
| `<base>_pagini_XX_YY.docx` | ambele | o livrare / bucată (interval `XX-YY` → label `XX_YY`; o pagină → `XX`) |
| `<base>_COMPLET_START_END.docx` | `telc-batch.bat` (via `uneste.py`) | documentul final unit pentru tot intervalul |
| `<base>_progres.txt` | `telc-batch.bat` | evidența bucăților terminate (vezi §6) |

> **Label-ul de pagini:** `-` din interval devine `_` în numele fișierului.
> `14-19` → `..._pagini_14_19.docx`. O singură pagină `18` → `..._pagini_18.docx`.

---

## 6. Reluarea la `telc-batch.bat` (progres + limită de tokeni)

### Fișierul de progres
`Z:\<base>_progres.txt` conține câte o linie per bucată terminată (ex. `28-29`).
O bucată e considerată **gata doar dacă** linia ei e în progres **ȘI** documentul `.docx` chiar există
pe disc. Astfel o bucată coruptă/ștearsă se reface automat la reluare.

### Cum reiau
**Reia exact aceeași comandă.** Scriptul sare peste bucățile deja gata și continuă de unde a rămas:
```bat
telc-batch.bat "telc_b2.pdf" 10-50    :: rulat din nou -> continua, nu o ia de la capat
```

### Dacă lovesc limita de tokeni
Scriptul detectează în ieșirea Claude mesaje de tip *„session limit" / „usage limit" / „resets"*,
se oprește **curat** la bucata curentă (cod de ieșire 2) și afișează un mesaj. Bucățile anterioare
rămân salvate în `Z:\`. **După resetarea limitei, reia aceeași comandă** — continuă de la bucata la
care s-a oprit. Documentul unit se creează automat la rularea care termină toate bucățile.

---

## 7. Unelte folosite în culise

- **`extrage.py INTERVAL [CALE_PDF]`** — extrage din PDF, pentru fiecare pagină, `pagina_NN.txt`
  (text brut) + `pagina_NN.png` (imagine 150 dpi) în `pagini_extrase/`. Numere PDF 1-based.
  Fără `CALE_PDF`, caută un PDF în folderul curent (preferă numele cu „telc").
- **`uneste.py "<base>" START END`** — unește bucățile intervalului, în ordinea paginilor, în
  `<base>_COMPLET_START_END.docx` (cu `docxcompose`). Inserează un paragraf gol între bucăți ca
  tabelele să nu fuzioneze. Iese cu eroare dacă lipsește vreo bucată din interval.
- **`fix.py FISIER.docx`** / **`valid.py FISIER.docx`** — fix XML obligatoriu (`tblInd`) + validare.

---

## 8. Depanare — erori tipice

| Mesaj / simptom | Cod ieșire | Ce înseamnă & ce fac |
|---|---|---|
| `EROARE: PDF-ul nu exista in Z:\` | 1 | Numele e greșit sau PDF-ul nu e în `Z:\`. Verifică numele exact (fără diacritice) și locația. |
| `Limita atinsa la bucata ...` | 2 (`telc-batch`) | Limită de tokeni. Bucățile dinainte sunt salvate. Reia **aceeași comandă** după resetare. |
| `Esec la bucata ... (documentul nu a fost generat; nu pare limita)` | 3 (`telc-batch`) | Apelul Claude nu a produs `.docx`. Cea mai frecventă cauză: **nume PDF cu diacritice** → redenumește ASCII. Vezi și ieșirea + `%TEMP%\telc_batch_out.txt` și promptul din `_temp\prompt_<label>.txt`. Reia comanda. |
| `EROARE la unire` | 4 (`telc-batch`) | `uneste.py` nu a putut combina bucățile (lipsește o bucată din interval). Bucățile rămân în `Z:\`; reia comanda ca să regenereze ce lipsește și să unească. |
| `documentul nu a fost generat la ...` | 1 (`telc.bat`) | Apelul Claude nu a salvat ieșirea. Vezi ieșirea de mai sus; verifică numele PDF (diacritice) și reia. |
| Doar prima bucată se procesează, restul nu | — | Ai rulat batch-ul **din interiorul unei sesiuni Claude**. Rulează-l dintr-un terminal normal (vezi §1). |
| `EROARE: PyMuPDF nu este instalat` (de la `extrage.py`) | — | `pip install pymupdf`. |
| `lipseste docxcompose` (la unire) | — | `pip install docxcompose` (`telc-batch.bat` încearcă să-l instaleze singur). |

---

## 9. Flux tipic recomandat

```bat
:: 1. asigura-te ca PDF-ul e in Z:\ cu nume ASCII
::    Z:\telc_b2.pdf

:: 2. o livrare rapida de test (o pagina)
telc.bat "telc_b2.pdf" 28

:: 3. multumit? proceseaza un interval mare, reluabil
telc-batch.bat "telc_b2.pdf" 10-50
::    daca lovesti limita: asteapta resetarea si reia exact aceeasi comanda
::    -> rezultat final: Z:\telc_b2_COMPLET_10_50.docx
```

---

## 10. Recuperare dupa reinstalare completa (de la zero)

Daca pierzi VM-ul / Windows-ul si trebuie sa reconstruiesti tot sistemul de la zero, urmeaza pasii de mai jos in ordine. La final ar trebui sa ai exact acelasi flux functional (`telc.bat` / `telc-batch.bat`).

### 1. Reinstaleaza uneltele de baza

Instaleaza (de pe site-urile oficiale / installer Windows):

- **Node.js LTS** (include `npm`)
- **Git**
- **Python 3** (bifeaza „Add Python to PATH" la instalare; aduce `pip`)
- **Claude Code** (CLI-ul `claude`)

Verifica fiecare:
```bat
node -v
npm -v
git --version
python --version
pip --version
claude --version
```

### 2. Cloneaza repo-ul de pe GitHub

```bat
git clone https://github.com/cristiv123/telc-traducere-bilingva.git
cd telc-traducere-bilingva
```
Asta aduce tot ce e versionat: **scripturile** (`extrage.py`, `uneste.py`, `fix.py`, `valid.py`, `telc.bat`, `telc-batch.bat`), **documentatia** (acest `MANUAL_SCRIPTURI.md`, `CLAUDE.md`) si **copia versionata a skill-ului** (in subfolderul `skill/`, fisierul `skill/SKILL.md`). PDF-ul sursa NU e in repo (vezi pasul 5).

### 3. Reinstaleaza skill-ul unde il citeste Claude Code

Claude Code citeste skill-ul din `C:\Users\<utilizator>\.claude\skills\telc-bilingual-docx\SKILL.md`, NU din repo. Copiaza-l acolo (creand folderul daca nu exista). Inlocuieste `<utilizator>` cu numele tau de utilizator Windows:

```bat
mkdir "C:\Users\<utilizator>\.claude\skills\telc-bilingual-docx"
copy skill\SKILL.md "C:\Users\<utilizator>\.claude\skills\telc-bilingual-docx\SKILL.md"
```
(In PowerShell: `New-Item -ItemType Directory -Force "C:\Users\<utilizator>\.claude\skills\telc-bilingual-docx"` apoi `Copy-Item skill\SKILL.md "C:\Users\<utilizator>\.claude\skills\telc-bilingual-docx\SKILL.md"`.)

### 4. Instaleaza dependentele

**Python** (folosite de `extrage.py` si `uneste.py`; `fix.py`/`valid.py` folosesc doar biblioteci standard):
```bat
pip install pymupdf docxcompose python-docx
```
- `pymupdf` → modulul `fitz` din `extrage.py` (extragere text + imagini din PDF).
- `docxcompose` + `python-docx` → `uneste.py` (unirea bucatilor intr-un singur docx).

**Node** — biblioteca `docx` se instaleaza **global** (scripturile de generare ruleaza cu `NODE_PATH` catre `npm root -g`):
```bat
npm install -g docx
```

### 5. Reconfigureaza si verifica

- **PDF-ul sursa nu e in repo** (e prea mare; `.gitignore` exclude `*.pdf`). Pune-l manual in `Z:\`, cu **nume ASCII** (fara diacritice), ex. `Z:\telc_b2.pdf`. Ambele batch-uri cauta sursa in `Z:\` si scriu rezultatele tot in `Z:\`.
- Verifica ca tot fluxul ruleaza, dintr-un **terminal normal** (cmd/PowerShell, NU din interiorul unei sesiuni Claude — vezi §1), cu o livrare mica de test:
  ```bat
  telc.bat "telc_b2.pdf" 28
  ```
  Daca genereaza `Z:\telc_b2_pagini_28.docx` si il deschide in LibreOffice, sistemul e refacut. Pentru un interval reluabil: `telc-batch.bat "telc_b2.pdf" 10-50`.

### 6. NOTA — sincronizarea skill-ului (doua locuri!)

Skill-ul exista in **doua** locatii:
- **copia de lucru** pe care o citeste Claude Code: `C:\Users\<utilizator>\.claude\skills\telc-bilingual-docx\SKILL.md`
- **copia versionata** din repo: `skill\SKILL.md`

`git clone` aduce DOAR copia versionata din repo. Invers, modificarile pe care le faci la copia de lucru NU ajung automat in repo.

> **Dupa ORICE modificare a skill-ului**, copiaza-l inapoi in `skill\SKILL.md` si fa commit + push, altfel GitHub ramane cu o versiune veche si urmatoarea recuperare (pasul 3) reinstaleaza un skill invechit:
> ```bat
> copy "C:\Users\<utilizator>\.claude\skills\telc-bilingual-docx\SKILL.md" skill\SKILL.md
> git add skill\SKILL.md
> git commit -m "Actualizare skill telc-bilingual-docx"
> git push
> ```
