@echo off
REM ===========================================================
REM SCRIPT: Creazione di una cartella con controllo esistenza
REM ===========================================================

@REM 1) Dichiarazione della variabile di ambiente TARGET
@REM    - TARGET è un nome che scegliamo noi (poteva chiamarsi PERCORSO o DESTINAZIONE)
@REM    - Contiene il percorso della cartella che vogliamo controllare o creare
@REM    - Quando scriviamo %TARGET%, Windows sostituisce questo segnaposto
@REM      con il contenuto effettivo della variabile
@REM  ATTENZIONE, ho inserito il nome MSys_TB02_TIPO_OGGETTI3 a posto di MSys_TB02_TIPO_OGGETTI per evitare
@REM              che la cartella originale venisse sovrascritta per cui nella ricostruzione    
@REM              della path originale occorre togliore l'estenzione.   
set "TARGET=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB02_TIPO_OGGETTI3\"

REM 2) Controllo se la cartella già esiste
REM    - Il comando IF EXIST verifica se l’oggetto (file o cartella) esiste nel percorso
if exist "%TARGET%" (
    REM Caso A: la cartella esiste già → mostro un messaggio
    echo La cartella esiste già: "%TARGET%"
) else (
    REM Caso B: la cartella non esiste → la creo con mkdir
    mkdir "%TARGET%"
    echo Cartella creata con successo: "%TARGET%"
)

REM 3) Pausa finale per permettere all’utente di leggere i messaggi
pause
