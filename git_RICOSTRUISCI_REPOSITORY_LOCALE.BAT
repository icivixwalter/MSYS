@echo off
REM ===============================================================
REM BATCH PER INIZIALIZZARE/SINCRONIZZARE LA REPOSITORY LOCALE GIT
REM ===============================================================
REM ESEGUI QUESTO BATCH DENTRO LA CARTELLA MSYS (NON CREA ALTRE CARTELLE)
REM ===============================================================

REM ---------------------------
REM 1) Controllo se la cartella è già una repo Git
REM ---------------------------
git rev-parse --is-inside-work-tree >NUL 2>&1
if ERRORLEVEL 1 (
    echo Inizializzo repository Git nella cartella corrente...
    git init
) else (
    echo Questa cartella e' gia' una repository Git.
)

REM ---------------------------
REM 2) Imposto o aggiorno il remote 'origin'
REM ---------------------------
git remote | findstr /I "^origin$" >NUL 2>&1
if ERRORLEVEL 1 (
    echo Aggiungo remote origin...
    git remote add origin https://github.com/icivixwalter/MSYS.git
) else (
    echo Aggiorno URL remote origin...
    git remote set-url origin https://github.com/icivixwalter/MSYS.git
)

REM ---------------------------
REM 3) Recupero dati dal remoto
REM ---------------------------
echo Recupero dati dal remoto...
git fetch origin

REM ---------------------------
REM 4) Gestione branch locale in base a branch remoto disponibile
REM Prova in ordine: calcolo -> master -> main
REM ---------------------------
set "BRANCH_FOUND="

for %%BR in calcolo master main do (
    set "REF="
    for /f "tokens=*" %%A in ('git ls-remote --heads origin %%BR') do set "REF=%%A"
    if defined REF (
        set "BRANCH_FOUND=%%BR"
        goto :CHECKOUT_BRANCH
    )
)

:CHECKOUT_BRANCH
if defined BRANCH_FOUND (
    echo Branch remoto '%BRANCH_FOUND%' trovato. Creo/aggiorno branch locale...
    git checkout -B %BRANCH_FOUND% origin/%BRANCH_FOUND%
    git branch --set-upstream-to=origin/%BRANCH_FOUND%_
