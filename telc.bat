@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

REM ============================================================
REM  telc.bat - genereaza un document bilingv DE/RO dintr-o
REM  singura comanda, neinteractiv, prin Claude Code.
REM
REM  Utilizare:
REM    telc.bat "nume_fisier.pdf" 14-19
REM    telc.bat "nume_fisier.pdf" 18
REM
REM  arg1 = numele PDF-ului sursa, aflat in Z:\
REM  arg2 = interval de pagini (14-19) sau o pagina (18)
REM
REM  Flag-uri OPTIONALE (dupa arg2, in orice ordine):
REM    --solutii "cheie.pdf"   al doilea PDF (cheia), aflat tot in Z:\
REM    --lectii  L3            lectiile asteptate (L3 sau L3,L4); OBLIGATORIU cu --solutii
REM    --dry-run               afiseaza promptul in loc sa cheme claude (pentru test)
REM
REM  Fara --solutii: comportament identic cu inainte (doar exercitii, fara solutii).
REM  Cu --solutii: pentru fiecare exercitiu se ataseaza solutia corelata pe eticheta
REM  exacta din cheie, sub "Lektion N" (vezi skill-ul telc-bilingual-docx).
REM ============================================================

set "PROJ=C:\claude-lab\02-traducere-telc"
set "SOFFICE=C:\Program Files\LibreOffice\program\soffice.exe"

REM --- validare: argumente prezente ---
if "%~1"=="" goto :usage
if "%~2"=="" goto :usage

set "PDFNAME=%~1"
set "PAGES=%~2"
set "SRC=Z:\%PDFNAME%"

REM --- validare: arg1 trebuie sa se termine in .pdf ---
echo(%PDFNAME%| findstr /i /e ".pdf" >nul
if errorlevel 1 goto :usage

REM --- validare: arg2 trebuie sa fie N sau N-M ---
echo(%PAGES%| findstr /r /x "[0-9][0-9]* [0-9][0-9]*-[0-9][0-9]*" >nul
if errorlevel 1 goto :usage

REM --- parsare flag-uri optionale dupa arg2 (--solutii / --lectii / --dry-run) ---
set "SOLPDF="
set "LECTII="
set "DRYRUN=0"
shift
shift
:parseflags
if "%~1"=="" goto :doneflags
if /i "%~1"=="--solutii" (
  set "SOLPDF=%~2"
  shift
  shift
  goto :parseflags
)
if /i "%~1"=="--lectii" (
  set "LECTII=%~2"
  shift
  shift
  goto :parseflags
)
if /i "%~1"=="--dry-run" (
  set "DRYRUN=1"
  shift
  goto :parseflags
)
echo.
echo EROARE: argument necunoscut: %~1
goto :usage
:doneflags

REM --- nume de iesire derivat din sursa + pagini ---
REM BASE = numele fara extensia .pdf (din PDFNAME, fiindca %1 a fost deja shift-at de flag-uri)
for %%F in ("%PDFNAME%") do set "BASE=%%~nF"
REM PAGESLABEL = pagini cu '-' inlocuit de '_'
set "PAGESLABEL=%PAGES:-=_%"
set "OUT=Z:\%BASE%_pagini_%PAGESLABEL%.docx"

echo.
echo === telc.bat ===
call :log "Pornire - sursa: %PDFNAME% - pagini: %PAGES%"

REM --- verificare OBLIGATORIE a existentei PDF-ului in Z:\ ---
call :log "Verific existenta PDF-ului in Z:\ ..."
if not exist "%SRC%" (
  echo.
  echo EROARE: PDF-ul nu exista in Z:\
  echo   Cautat: "%SRC%"
  echo Verifica numele fisierului si incearca din nou.
  echo.
  exit /b 1
)
call :log "PDF gasit: %SRC%"
call :log "Iesire planificata: %OUT%"

cd /d "%PROJ%"

REM --- mod SOLUTII: validare cheie + garda lectii + maparea lectie-pagini de cheie.
REM     Flux PLAT (goto, nu un bloc imbricat) ca 'exit /b N' sa propage corect codul. ---
set "SOLSRC="
set "SOLMAP="
if not defined SOLPDF goto :after_solutii

if not defined LECTII (
  echo.
  echo EROARE: --solutii necesita si --lectii ^(ex. --lectii L3^).
  goto :usage
)
set "SOLSRC=Z:\%SOLPDF%"
if not exist "!SOLSRC!" (
  echo.
  echo EROARE: PDF-ul de solutii nu exista in Z:\: "!SOLSRC!"
  exit /b 1
)
call :log "Cheie solutii: !SOLSRC! - lectii asteptate: %LECTII%"
python lectii.py verify "%SRC%" %PAGES% %LECTII%
if errorlevel 1 (
  call :log "Oprire: lectiile detectate nu corespund cu --lectii (vezi mesajul de mai sus)."
  exit /b 5
)
for /f "delims=" %%s in ('python lectii.py plan "!SOLSRC!" %LECTII%') do set "SOLMAP=%%s"
if not defined SOLMAP (
  call :log "Oprire: nu am putut determina paginile de cheie pentru %LECTII%."
  exit /b 5
)
call :log "Mapare solutii: !SOLMAP!"
:after_solutii

REM --- instructiunea de baza pentru Claude Code (neinteractiv) ---
REM     (ramura fara --solutii e identica cu varianta veche: extragere in pagini_extrase/)
set "PROMPT=Foloseste skill-ul telc-bilingual-docx. Genereaza un document Word bilingv germana-romana pentru paginile %PAGES% din PDF-ul aflat la '%SRC%'. Pasi: 1) extrage paginile %PAGES% ruland comanda: python extrage.py %PAGES% '%SRC%' (text + imagini in pagini_extrase/); 2) compara OBLIGATORIU textul OCR cu imaginile pagina_NN.png pentru a corecta greselile de OCR; 3) tradu integral, propozitie cu propozitie, in tabel bilingv DE/RO conform skill-ului (fara diacritice in coloana romana, caractere germane corecte, bold pe termeni); 4) aplica fix-ul XML; 5) salveaza documentul .docx final exact la calea '%OUT%'. Anunta pe scurt fiecare pas pe masura ce il faci: 'Extrag paginile %PAGES%', 'Corectez OCR comparand cu imaginile', 'Traduc', 'Aplic fix XML', 'Salvez la %OUT%'. Nu cere confirmari, ruleaza pana la capat."

REM --- augmentare cu solutii (doar daca --solutii a fost dat) ---
if defined SOLPDF set "PROMPT=!PROMPT! SOLUTII (vezi sectiunea 'doua surse' din skill): cheia de raspunsuri e in PDF-ul separat '!SOLSRC!'. Maparea lectie-pagini de cheie: !SOLMAP!. Pentru fiecare exercitiu de pe paginile de exercitii: determina lectia din titlul 'Lektion N' (confirmat pe imagine), extrage paginile de cheie ale ACELEI lectii cu python extrage.py PAGINILE_DE_CHEIE '!SOLSRC!' --sub loes, si adauga la finalul tabelului exercitiului o sectiune de solutii cu eticheta 'SOLUTII / LOESUNGEN - Lektion N / Xy (din cheie: Lektion N / Xy)' si raspunsul corelat pe ETICHETA EXACTA (exercitiul Xy ia solutia Xy de sub 'Lektion N' in cheie). Corecteaza cheia compacta pe imagine (e sistematic corupta de OCR). Reda FIDEL anomaliile de tipar din cheie: scrie exact cum e tiparit si adauga un rand de avertizare cu completarea logica, NU repara tacit."

if "%DRYRUN%"=="1" (
  echo.
  echo === DRY-RUN: prompt care AR FI trimis lui claude ^(nu se executa^) ===
  echo.
  echo !PROMPT!
  echo.
  echo === sfarsit dry-run ===
  endlocal
  exit /b 0
)

call :log "Extragere si traducere in curs (poate dura cateva minute)..."
echo.
claude --dangerously-skip-permissions -p "!PROMPT!"
echo.

REM --- deschide rezultatul daca a fost generat ---
if exist "%OUT%" (
  call :log "Document generat: %OUT%"
  call :log "Deschid in LibreOffice..."
  start "" "%SOFFICE%" "%OUT%"
) else (
  call :log "EROARE: documentul nu a fost generat la %OUT%."
  echo Verifica iesirea de mai sus.
  exit /b 1
)

endlocal
exit /b 0

REM ============================================================
:log
echo [%TIME%] %~1
exit /b 0

REM ============================================================
:usage
echo.
echo telc.bat - genereaza un document Word bilingv DE/RO dintr-o
echo singura comanda, din paginile unui PDF aflat in Z:\
echo.
echo Sintaxa:
echo   telc.bat "nume.pdf" ^<pagina sau interval^> [--solutii "cheie.pdf" --lectii L3] [--dry-run]
echo.
echo Exemple:
echo   telc.bat "telc_b2.pdf" 18        (o singura pagina)
echo   telc.bat "telc_b2.pdf" 14-19     (un interval de pagini)
echo   telc.bat "Sicher_C1_1_Arbeitsbuch_ocr.pdf" 35-50 --solutii "Sicher_C1_1_AB_Loesungen_ocr.pdf" --lectii L3
echo.
echo Flag-uri optionale (dupa interval):
echo   --solutii "cheie.pdf"  ataseaza solutiile corelate din cheie (al doilea PDF din Z:\)
echo   --lectii  L3           lectiile asteptate (L3 sau L3,L4); obligatoriu cu --solutii
echo   --dry-run              afiseaza promptul in loc sa cheme claude (pentru test)
echo.
exit /b 1
