
:------------------------------DOS apri file con DISCO + PATH + FILE - imposto le variabili iniziali
@echo off
@echo off



@REM attenzione deve aprire solo con word 2000 non 2024 e la path del word 2000
@REM è questa : ???  crearea con AI  un comando che apre solo con word 2000???
@REM la path dell'exe si trova 	qui: C:\Program Files (x86)\Microsoft Office\Office
@REM l'exe si trova 						qui: WINWORD.exe

@REM 1) Percorso dell'eseguibile di Word 2000: Ho aggiunto la variabile WINWORD2000 che contiene il percorso completo dell'eseguibile 
@REM di Word 2000.
@REM 2) Comando START: Ho modificato i comandi START per aprire il file specificando esplicitamente il percorso di WINWORD2000. 
@REM In questo modo, il file .doc verrà aperto con Word 2000
@REM 3) Uso di explorer.exe per aprire la cartella: Ho aggiunto explorer.exe al comando che apre la cartella, 
@REM così da evitare conflitti con il comando di chiusura (exit).




@echo off

:------------------------------- CONTROLLO VARIABILI SET + APERTURA + APRO SOLO CON WORD 2000
echo.
echo CONTROLLO VARIABILI IMPOSTATE  
echo. 
echo "IL DISCO 	: ----> " %DISCO_s%
echo "LA PATH 	: ----> " %PATH_s%
echo "IL FILE   : ----> " %FILE_s%
echo.

REM Modifica per usare Word 2000
SET WINWORD2000="C:\Program Files (x86)\Microsoft Office\Office\WINWORD.exe"


SET DISCO_s=c:
SET PATH_s=\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\TUTORIAL\
SET FILE_s=GIORNALE_OPERAZIONI_PROGETTO_MSYS.doc

START "APRI FILE	" %WINWORD2000% "%DISCO_s%%PATH_s%%FILE_s%"
START "APRI CARTELLA	" explorer.exe "%DISCO_s%%PATH_s%" &^ exit











@GOTO note
  Attenzione ci sono due tipi di path quello con la path relativo o assoluta, secondo i due esempi 
  path relativa es. ---> "\CASA\PRES3000_07\WALTER_ATTI\WALTER_TIM_SCHEDA_PENNETTA_3662631843_(WIFI_ALCATEL).msg" &^ exit
  path assoluta es. ---> START "APRI" /i "c:\CASA\PRES3000_07\WALTER_ATTI\WALTER_TIM_SCHEDA_PENNETTA_3662631843_(WIFI_ALCATEL).msg" &^ exit
 
  la procedura di start attiva con la path relativa è valida solo per i disco c: mentre in caso di rete occorre quella assoluta.
  Inoltre questo esempio imposta un indirizzo assoluto con 3 variabili :
    			DISCO_s + PATH_s + FILE_s 
    			ATTENZIONE a non inserire spazi nel comando SET es. set DISCO_s=Y: (ok) set DISCO_s= Y: (NON FUNZIONA)
    				altrimenti non funziona
	ESEMPIO FUNZIONANTE
		Nella procedura di seguito indicata vengono impostate le 3 variabili stringhe DISCO, PATH E FILE	
		e vengono unite con la & commerciale senza virgolette ("") in questo modo se il file
		viene spostato in rete o su altro disco basta cambiare solo la variabile del disco
	
	ATTENZIONE
		non occorre start "..." /I PERCHE ALTRIMENTI NON APRE.
		Basta Start "...." 
	ATTENZIONE	
		non occorrono inserire gli spazi tra le variabili
		%DISCO_s%%PATH_s%%FILE_s% 	= OK
		%DISCO_s% %PATH_s% %FILE_s%	= NO NON FUNZIONA
	
	ATTENZIONE
		la chiusura del teminale solo alla fine es:
		START "ATTIVITA" ....
		START "II ATTIVITA"...
		START "APRI CARTELLA	" %DISCO_s%%PATH_s% &^ exit
		
	ATTENZIONE PATH RELATIVA O ASSOLUTA
		Se vuoi utilizzare la path relativa ed applicare il bat in ogni disco
		occorre impostare la variabile DISCO in questo modo : DISCO_s=.
		Nel caso di path assoluta con indicazione del disco : DISCO_s=C:
	

@:note



