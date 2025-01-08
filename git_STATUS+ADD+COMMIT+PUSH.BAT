@REM  cd.. devi posizionarti su git_desktop
CD c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\
dir 

@REM  fai un controllo con status, poi aggiungi con add tutto l'albero + il comimit e il push
git status

ECHO pausa di 3 secondi per il controllo
@TIMEOUT /T 3 /NOBREAK



git add .
git commit -m "aggiornamento oggetti MSYS"

@rem pausa di 20 secondi per consentire il commit

@REM 01)
@REM sospensione per 1 secondo con crononometro 
@REM ---->	timeout /t 2 /nobreak > NUL
@REM ---->
ECHO
ECHO
ECHO  pausa di 20 secondi per consentire il commit
ECHO
@TIMEOUT /T 20 /NOBREAK

git push


@rem pausa di  secondi per il controllo
ECHO

ECHO pausa di 3 secondi per il controllo
@TIMEOUT /T 3 /NOBREAK
