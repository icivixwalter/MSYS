@REM  cd.. devi posizionarti su git_desktop
CD c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\

@ECHO..						IMPEGNI	 *** INIZIO ***										
@ECHO...................................................................................................................
@ECHO..
@ECHO SALVATAGGIO ESTERNO SU DISCO LOCALE IMPEGNI
@ECHO..
START "SALVA ARCHIVIO IMPEGNI" CALL "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\IMPEGNI\ZIP_7ZIP_INCREMENTALE_{CRIPTA+DECODIFICA_INTERNO}.BAT"

@ECHO..
@ECHO SALVATAGGIO SU GITHUB DEL PROGETTO IMPEGNI MEDIANTE GIT PUSH..
@ECHO..
START "SALVA ARCHIVIO IMPEGNI" CALL "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\IMPEGNI\git_STATUS+ADD+COMMIT+PUSH.BAT"

@ECHO..						IMPEGNI  *** FINE ***											
@ECHO...................................................................................................................




@ECHO..						MENU	 *** INIZIO ***										
@ECHO...................................................................................................................

@ECHO..
@ECHO SALVATAGGIO ESTERNO SU DISCO LOCALE MENU
@ECHO..
START "SALVA ARCHIVIO IMPEGNI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MENU\ZIP_7ZIP_INCREMENTALE_{CRIPTA+DECODIFICA_INTERNO}.BAT"

@ECHO..
@ECHO SALVATAGGIO SU GITHUB DEL PROGETTO MENU MEDIANTE GIT PUSH..
@ECHO..
START "SALVA ARCHIVIO IMPEGNI" CALL "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\MENU\git_STATUS+ADD+COMMIT+PUSH.BAT"


@ECHO..						MENU	 *** FINE ***										
@ECHO...................................................................................................................




@ECHO..						MSYS	 *** INIZIO ***										
@ECHO...................................................................................................................

@ECHO..
@ECHO SALVATAGGIO ESTERNO SU DISCO LOCALE MSYS
@ECHO..
START "SALVA ARCHIVIO IMPEGNI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\ZIP_7ZIP_INCREMENTALE_{CRIPTA+DECODIFICA_INTERNO}.BAT"


@ECHO..
@ECHO SALVATAGGIO SU GITHUB DEL PROGETTO MSYS MEDIANTE GIT PUSH..
@ECHO..
START "SALVA ARCHIVIO IMPEGNI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\git_STATUS+ADD+COMMIT+PUSH.BAT"



@ECHO..						MSYS	 *** FINE  ***										
@ECHO...................................................................................................................


ECHO pausa di 3 secondi per il controllo
@TIMEOUT /T 3 /NOBREAK



git pull

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

