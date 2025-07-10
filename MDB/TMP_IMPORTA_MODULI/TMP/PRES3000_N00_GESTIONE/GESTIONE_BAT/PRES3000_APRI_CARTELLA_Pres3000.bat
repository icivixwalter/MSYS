
@ECHO OFF

@REM directory di salvataggio:
@REM ATTENZIONE la path deve essere avvolta tra "" per evitare errori
@SET PATH_s = c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N00_GESTIONE\





:------------------------------------APRO LA CARTELLA PRINCIPALE 
start "Apro la cartella PRES300" /B "%PATH_s%" & ^exit






















GOTO TUTORIAL

:________________________APRO LA CARTELLA PRINCIPALE 
GOTO:MEMO
	
	 //NOTA PATH RELATIVA ED ASSOLUTA:
	 //-----------------------------------------------------------------------------------------------------------//
		  Attenzione ci sono due tipi di path quello con la path relativo o assoluta, secondo i due esempi 
		  path relativa es. ---> "\CASA\PRES3000_07\WALTER_ATTI\WALTER_TIM_SCHEDA_PENNETTA_3662631843_(WIFI_ALCATEL).msg" &^ exit
		  path assoluta es. ---> START "APRI" /i "c:\CASA\PRES3000_07\WALTER_ATTI\WALTER_TIM_SCHEDA_PENNETTA_3662631843_(WIFI_ALCATEL).msg" &^ exit
	
		  la procedura di start attiva con la path relativa è valida solo per i disco c: mentre in caso di rete occorre quella assoluta.
		  Inoltre questo esempio imposta un indirizzo assoluto con 3 variabili :
					DISCO_s + PATH_s + FILE_s 
					ATTENZIONE a non inserire spazi nel comando SET es. set DISCO_s=Y: (ok) set DISCO_s= Y: (NON FUNZIONA)
						altrimenti non funziona
						
					ATTENZIONE PATH RELATIVA O ASSOLUTA
					Se vuoi utilizzare la path relativa ed applicare il bat in ogni disco
					occorre impostare la variabile DISCO in questo modo : DISCO_s=.
					Nel caso di path assoluta con indicazione del disco : DISCO_s=C:
				
		ESEMPIO FUNZIONANTE
			Nella procedura di seguito indicata vengono impostate le 3 variabili stringhe DISCO, PATH E FILE	
			e vengono unite con la & commerciale senza virgolette ("") in questo modo se il file
			viene spostato in rete o su altro disco basta cambiare solo la variabile del disco
				
				ESEMPIO SET CON PATH RELATIVA 
					@REM directory di salvataggio: ..\ = Due posizione compresa la posizione corrente
						@SET PATH_s=..\

	
				
		ATTENZIONE START
			non occorre start "..." /I PERCHE ALTRIMENTI NON APRE.
			Basta Start "...." 
		ATTENZIONE SET VARIABILI	
			non occorrono inserire gli spazi tra le variabili
			%DISCO_s%%PATH_s%%FILE_s% 	= OK
			%DISCO_s% %PATH_s% %FILE_s%	= NO NON FUNZIONA
		
		ATTENZIONE EXIT 
			la chiusura del teminale solo alla fine es:
			START "ATTIVITA" ....
			START "II ATTIVITA"...
			START "APRI CARTELLA	" %DISCO_s%%PATH_s% &^ exit
			
			ATTENZIONE UTILIZZO DI &^ exit PER LA CHIUSURA DELLA FINESTRA CMD CON LA BAT O CON START
				se viene aperto un file o una cartella con un richiamo ad un'altra procedera bat
				NON OCCORRE UTILIZZARE START ma occorre utiliccare CALLE come in questo esempio:
					::-------------------------------RICHIAMO UN BAT CON IL CALL
					CALL %DISCO_s%%PATH_s%%FILE_s%&^ exit
	
				Se utilizzo uno start allora la chiusura si applica in questo modo:
					START "APRI CARTELLA	" %DISCO_s%%PATH_s%&^ exit
	//-----------------------------------------------------------------------------------------------------------//

	
	//NOTA PER L'UTILIZZO DI START
	//-----------------------------------------------------------------------------------------------------------//
		Nel caso di apertura in successione utilizzare START
			es. Start "... nome" path &^exit
				utilizzare il comando * &^exit * intervallato da uno spazio per ogni comando start
				Non è necessario utilizzare l'opzione /B LA QUALE VIENE UTILIZZARE PER NON DUPLICARE IL PROCESSO
				altrimenti il cmd rimane aperto.
				Non è necessario utilizzare il comando CALL
					esempi:
						1) start "Apro IL DB 1" "c:\casa\file.mdb" &^exit 
						2) start "Apro IL DB collegato lnk" "c:\casa\file.lnk" &^exit 
						3) start "Apro la directory" "c:\casa\" &^exit 
	//-----------------------------------------------------------------------------------------------------------//					
							
:MEMO

:------------------------------DOS apri file con DISCO + PATH + FILE - imposto le variabili iniziali
@echo off

:------------------------------- IMPOSTO le variabili DISCO, PATH E FILE attenzione senza spazi dopo l'uguale
@REM attenzione nel caso che il valore del disco venga prelevato da un file esterno es. un txt occorre inserire
@rem il parametro /P = valore di input e il segno uguale e l'operatore di scorrimento sinistro < ma devono essere uniti 
@rem es. ... =<.. non staccati. -- altrimenti non viene caricato il valore esterno. 
@rem esempio SET /P DISCO_s=<DISCO_Default.TXT
	
	@REM va bene anche barra doppia e singola \\ o \
	:----------------------------------------DISCO + PATH + FILE
	SET /P DISCO_s=<MODELLO_BAT_START_Aprofile_DISCO_Default.TXT
	SET PATH_01_s=\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\
	SET PATH_02_s=\LLPP_ARCHIVI_MDB\AA_MODELLO\
	SET PATH_03_s=\MODELLO_TEX+TXT_ANALISI\
	SET FILE_s=MODELLO_ANALISI_TEX.tex
	
	
	



echo. "........................................................."
echo. "	   CONTROLLO PARAMETRI CON DEL COMANDO SET 		"

echo. path ricostruita: \GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\\LLPP_ARCHIVI_MDB\AA_MODELLO\\MODELLO_TEX+TXT_ANALISI\
echo. 
echo. "il disco caricato dal file variabile DISCO_s:"
echo  %DISCO_s%
echo. "la path  PATH_01_s :"
echo. %PATH_01_s%


echo. "la path  PATH_02_s :"
echo. %PATH_02_s%


echo. "la path  PATH_03_s :"
echo. %PATH_03_s%


echo. "il file FILE_s :"
echo. %FILE_s%

echo. "TUTTA LA PATH ED il file FILE_s :"
echo. "%DISCO_s%%PATH_01_s%%PATH_02_s%%PATH_03_s%%FILE_s%"


echo. "...................................................."
:________________________APRO IL DB + LATEX = memo per chiudere la finestra utilizzare /B  e alla fine &^exit 



start "Apro la cartella " "%DISCO_s%%PATH_01_s%%PATH_02_s%%PATH_03_s%"&^ exit

start "Apro IL DB MDB" "%DISCO_s%%PATH_01_s%%PATH_02_s%%PATH_03_s%%FILE_s%"






TUTORIAL:





