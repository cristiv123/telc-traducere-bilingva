@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

REM ============================================================
REM  telc-batch.bat - procesare RELUABILA a paginilor unui PDF,
REM  pe bucati de 2 pagini, cu reluare si unire finala.
REM
REM  Utilizare:
REM    telc-batch.bat "nume.pdf" 10-50     (un interval)
REM    telc-batch.bat "nume.pdf" 18        (o singura pagina)
REM    telc-batch.bat "nume.pdf"           (tot fisierul)
REM
REM  NOTA: foloseste un nume de PDF fara diacritice / caractere
REM  speciale (cmd.exe corupe diacriticele in nume). Ex: telc_b2.pdf
REM ============================================================

set "PROJ=C:\claude-lab\02-traducere-telc"
set "SOFFICE=C:\Program Files\LibreOffice\program\soffice.exe"
set "TMPOUT=%TEMP%\telc_batch_out.txt"

REM --- validare: arg1 prezent si .pdf ---
if "%~1"=="" goto :usage
echo(%~1| findstr /i /e ".pdf" >nul
if errorlevel 1 goto :usage

set "PDFNAME=%~1"
set "SRC=Z:\%PDFNAME%"
set "BASE=%~n1"
set "PROGRES=Z:\%BASE%_progres.txt"

REM --- validare: daca arg2 e dat, trebuie N sau N-M ---
set "ARG2=%~2"
if not "%ARG2%"=="" (
  echo(%ARG2%| findstr /r /x "[0-9][0-9]* [0-9][0-9]*-[0-9][0-9]*" >nul
  if errorlevel 1 goto :usage
)

REM --- verificare OBLIGATORIE a existentei PDF-ului in Z:\ ---
if not exist "%SRC%" (
  echo.
  echo EROARE: PDF-ul nu exista in Z:\
  echo   Cautat: "%SRC%"
  echo Verifica numele fisierului si incearca din nou.
  echo.
  exit /b 1
)

REM --- determinarea intervalului de lucru ---
set "WHOLE=0"
if "%ARG2%"=="" (
  set "WHOLE=1"
  set "START=1"
  set "END="
  for /f "delims=" %%n in ('python -c "import fitz,sys;print(fitz.open(sys.argv[1]).page_count)" "%SRC%"') do set "END=%%n"
  if "!END!"=="" (
    echo EROARE: nu am putut determina numarul de pagini al PDF-ului.
    exit /b 1
  )
) else (
  echo(%ARG2%| findstr /r /x "[0-9][0-9]*-[0-9][0-9]*" >nul
  if errorlevel 1 (
    set "START=%ARG2%"
    set "END=%ARG2%"
  ) else (
    for /f "tokens=1,2 delims=-" %%a in ("%ARG2%") do (
      set "START=%%a"
      set "END=%%b"
    )
  )
)

if !START! GTR !END! (
  echo EROARE: interval invalid: !START! mai mare decat !END!.
  exit /b 1
)

set /a SPAN=END-START+1
set /a TOTAL=(SPAN+1)/2

echo.
echo === telc-batch.bat ===
call :log "Sursa: %PDFNAME% - interval: !START!-!END! - bucati: !TOTAL!"

REM --- confirmare DOAR pentru tot fisierul ---
if "%WHOLE%"=="1" (
  echo.
  echo !END! pagini -^> !TOTAL! bucati. Procesez pe rand; daca lovesc limita ma opresc curat si pot relua.
  echo.
  pause
)

cd /d "%PROJ%"

set "DONE=0"
set "K=0"
set "LASTOK="

REM ============================================================
REM  BUCLA PRINCIPALA - bucla goto cu corp INLINE (claude ruleaza
REM  la nivel de top, NU intr-o subrutina apelata - altfel cmd
REM  pierde pozitia dupa programul extern).
REM ============================================================
set "I=!START!"

:chunk_loop
if !I! GTR !END! goto :all_chunks_done

set "LO=!I!"
set /a HI=!I!+1
if !HI! GTR !END! set "HI=!END!"
if !LO! EQU !HI! ( set "PG=!LO!" ) else ( set "PG=!LO!-!HI!" )
set "LBL=!PG:-=_!"
set "OUTCHUNK=Z:\%BASE%_pagini_!LBL!.docx"
set /a K+=1
set "CURCHUNK=!PG!"

REM --- reluare: bucata e facuta DOAR daca e in progres SI fisierul exista ---
set "ISDONE="
if exist "%PROGRES%" (
  findstr /x /c:"!PG!" "%PROGRES%" >nul 2>&1 && if exist "!OUTCHUNK!" set "ISDONE=1"
)
if defined ISDONE (
  echo [!K!/!TOTAL!] paginile !PG! deja facute - sar peste.
  set /a DONE+=1
  set "LASTOK=!PG!"
  goto :next_chunk
)

echo.
echo [!K!/!TOTAL!] Procesez paginile !PG! ...

REM --- promptul se scrie INTR-UN FISIER (nu intr-o variabila lunga cu caractere
REM     speciale sub delayed expansion, care iesea goala si facea claude sa cada tacit).
REM     Redirect-ul e pus in fata; !PG!/!OUTCHUNK! se expandeaza, %SRC% la fel; promptul
REM     nu contine & | < > ^ deci e sigur de echo-at. ---
set "PROMPTFILE=%PROJ%\_temp\prompt_!LBL!.txt"
> "!PROMPTFILE!" echo Foloseste skill-ul telc-bilingual-docx. Genereaza un document Word bilingv germana-romana pentru paginile !PG! din PDF-ul aflat la '%SRC%'. Pasi: 1) extrage paginile !PG! ruland: python extrage.py !PG! '%SRC%' (text + imagini in pagini_extrase/); 2) compara OBLIGATORIU textul OCR cu imaginile pagina_NN.png pentru a corecta greselile de OCR; 3) tradu integral, propozitie cu propozitie, in tabel bilingv DE/RO conform skill-ului (fara diacritice in coloana romana, caractere germane corecte, bold pe termeni); 4) aplica fix-ul XML; 5) salveaza documentul .docx final exact la calea '!OUTCHUNK!'. Anunta pe scurt fiecare pas: 'Extrag paginile !PG!', 'Corectez OCR comparand cu imaginile', 'Traduc', 'Aplic fix XML', 'Salvez'. Nu cere confirmari, ruleaza pana la capat.

REM --- claude citeste promptul din STDIN (validat ca functioneaza) ---
type "!PROMPTFILE!" | claude --dangerously-skip-permissions -p > "%TMPOUT%" 2>&1
set "CLAUDE_RC=!errorlevel!"
type "%TMPOUT%"

REM --- mesaj clar daca apelul claude a esuat (sa nu mai cada tacit) ---
if not "!CLAUDE_RC!"=="0" (
  echo.
  call :log "EROARE: apelul claude a esuat (cod !CLAUDE_RC!) la bucata !PG!."
  echo Vezi iesirea de mai sus si "%TMPOUT%". Daca numele PDF are diacritice, redenumeste-l ASCII.
  echo Promptul folosit a fost salvat in: "!PROMPTFILE!"
)

REM --- detectare limita ---
findstr /i /c:"session limit" /c:"usage limit" /c:"resets" "%TMPOUT%" >nul
if not errorlevel 1 goto :limit

REM --- succes? documentul a fost generat ---
if not exist "!OUTCHUNK!" goto :failchunk

findstr /x /c:"!PG!" "%PROGRES%" >nul 2>&1 || >>"%PROGRES%" echo !PG!
echo OK: %BASE%_pagini_!LBL!.docx
set /a DONE+=1
set "LASTOK=!PG!"

:next_chunk
set /a I+=2
goto :chunk_loop

:all_chunks_done

REM ============================================================
REM  UNIRE FINALA (toate bucatile OK)
REM ============================================================
echo.
if "%TOTAL%"=="1" (
  set "FINAL=!OUTCHUNK!"
  call :log "O singura bucata; documentul final este chiar aceasta."
) else (
  call :log "Toate bucatile gata. Unesc documentele..."
  python -c "import docxcompose" 2>nul
  if errorlevel 1 (
    call :log "Instalez docxcompose..."
    pip install docxcompose >nul 2>&1
  )
  set "FINAL="
  for /f "delims=" %%f in ('python uneste.py "Z:\%BASE%" %START% %END%') do set "FINAL=%%f"
  if "!FINAL!"=="" (
    call :log "EROARE la unire. Bucatile raman in Z:\; reia comanda."
    exit /b 4
  )
  python fix.py "!FINAL!" >nul 2>&1
)

echo.
call :log "Gata. !DONE!/!TOTAL! bucati terminate. Ultima: !LASTOK!."
call :log "Document final: !FINAL!"
call :log "Deschid documentul final in LibreOffice..."
start "" "%SOFFICE%" "!FINAL!"

endlocal
exit /b 0

REM ============================================================
:limit
echo.
call :log "Limita atinsa la bucata %CURCHUNK%. Bucatile anterioare sunt salvate in Z:."
echo Reia aceeasi comanda dupa resetare ca sa continui de aici.
echo Documentul unit se va crea automat la rularea care termina toate bucatile.
exit /b 2

REM ============================================================
:failchunk
echo.
call :log "Esec la bucata %CURCHUNK% (documentul nu a fost generat; nu pare limita)."
echo Daca numele PDF-ului are diacritice/caractere speciale, redenumeste-l ASCII (ex. telc_b2.pdf).
echo Verifica iesirea de mai sus si reia comanda.
exit /b 3

REM ============================================================
:log
echo [%TIME%] %~1
exit /b 0

REM ============================================================
:usage
echo.
echo telc-batch.bat - proceseaza paginile unui PDF in bucati de 2 pagini,
echo cu reluare la limita de tokeni si unire finala intr-un singur document.
echo Sursa PDF se afla in Z:\, iar rezultatele se salveaza tot in Z:\.
echo.
echo Sintaxa:
echo   telc-batch.bat "nume.pdf" N-M     (un interval, ex. 10-50)
echo   telc-batch.bat "nume.pdf" N       (o singura pagina, ex. 18)
echo   telc-batch.bat "nume.pdf"         (tot fisierul)
echo.
echo Exemple:
echo   telc-batch.bat "telc_b2.pdf" 18
echo   telc-batch.bat "telc_b2.pdf" 10-50
echo   telc-batch.bat "telc_b2.pdf"
echo.
echo NOTA: foloseste un nume de PDF fara diacritice / caractere speciale.
echo.
exit /b 1
