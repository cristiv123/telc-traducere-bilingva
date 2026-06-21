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

REM --- nume de iesire derivat din sursa + pagini ---
REM BASE = numele fara extensia .pdf
set "BASE=%~n1"
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

REM --- instructiunea pentru Claude Code (neinteractiv) ---
set "PROMPT=Foloseste skill-ul telc-bilingual-docx. Genereaza un document Word bilingv germana-romana pentru paginile %PAGES% din PDF-ul aflat la '%SRC%'. Pasi: 1) extrage paginile %PAGES% ruland comanda: python extrage.py %PAGES% '%SRC%' (text + imagini in pagini_extrase/); 2) compara OBLIGATORIU textul OCR cu imaginile pagina_NN.png pentru a corecta greselile de OCR; 3) tradu integral, propozitie cu propozitie, in tabel bilingv DE/RO conform skill-ului (fara diacritice in coloana romana, caractere germane corecte, bold pe termeni); 4) aplica fix-ul XML; 5) salveaza documentul .docx final exact la calea '%OUT%'. Anunta pe scurt fiecare pas pe masura ce il faci: 'Extrag paginile %PAGES%', 'Corectez OCR comparand cu imaginile', 'Traduc', 'Aplic fix XML', 'Salvez la %OUT%'. Nu cere confirmari, ruleaza pana la capat."

call :log "Extragere si traducere in curs (poate dura cateva minute)..."
echo.
claude --dangerously-skip-permissions -p "%PROMPT%"
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
echo   telc.bat "nume.pdf" ^<pagina sau interval^>
echo.
echo Exemple:
echo   telc.bat "telc_b2.pdf" 18        (o singura pagina)
echo   telc.bat "telc_b2.pdf" 14-19     (un interval de pagini)
echo.
exit /b 1
