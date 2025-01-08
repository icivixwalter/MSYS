@REM  ***************************************************************

@REM      file batch : Salva-C rar
@REM      directory = c:\casa\LTT ; \DOC  e \STAMPE
@REM      file WinRAR

@REM   ***************************************************************


@REM 					LE SOSTITUZIONI GENERALI
@REM @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

@REM  		Disco e codice Disco
@REM .......................................................
@REM -----> 	C:
@REM -----> 	_C
@REM .......................................................


@REM		directory y dove archiviare i dati
@REM .......................................................
@REM ----->	c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\HELP\
@REM .......................................................


@REM 		Il nome di tutti i file di archivio dei dati
@REM .......................................................
@REM ----->	CorsoLazio_ScA_


@REM 		Il file di archivio dei dati
@REM .......................................................
@REM ----->	HELP_ANALISI_MSYS.docx


@REM 					LE SOSTITUZIONI GENERALI *** FINE ***
@REM @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@




ECHO OFF
CLS


@REM 			SALVATAGGIO DATI
REM *************************************************************************************************************************

@REM			MESSAGGIO OPERAZIONE
@REM -----------------------------------------------------------------------------------------
echo  "============================================================="
echo.				INIZIO
echo.	MSG - SALVATAGGIO DATI - HELP_ANALISI_MSYS.docx
echo.  
echo.
echo  "============================================================="

@REM			MESSAGGIO OPERAZIONE *** FINE ***
@REM -----------------------------------------------------------------------------------------







@REM			OPERAZIONE DI APERTURA FILE
@REM -----------------------------------------------------------------------------------------


@REM 02) Posso così caricare il mio valore di configurazione dal file root.config nella mia variabile root:
	@REM  digito il comando sotto indicato nel prompt di DOS. Con il comando Set viene
	@REM  impostata la variabile path del file da aprire e poi richiamata. CARICO LA PATH DA FILE ESTERNO
		
		
		
		@rem sospeso perche la nel file è diversa 
		SET DISCO_s=c:
		
		@rem la path è impostata manualmente
		SET PATH_s=\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\HELP\HELP_FORM\
		SET FILE_s=HELP_Msys_DF11_PROGETTI_Frm00_M01_GESTIONE.pdf
		echo off
		echo la path %DISCO_s% %path_s% ed il file da aprire %FILE_s%
		
		  START "APRO IL FILE" "%DISCO_s%%PATH_s%%FILE_s%"&^EXIT
		
		 	

@REM sospensione con conteggio secondi visibile, serve per riportare il controllo al presente comando
timeout /t 1 /nobreak 




@REM			OPERAZIONE DI APERTURA FILE *** FINE ***
@REM -----------------------------------------------------------------------------------------




@REM			SOSPENSIONE
@REM -----------------------------------------------------------------------------------------

	@REM 01)
	@REM sospensione per 1 secondo con crononometro 
	@REM timeout /t 2 /nobreak > NUL
	@REM ---->	TIMEOUT /T 6 /NOBREAK


@REM			SOSPENSIONE *** FINE ***
@REM -----------------------------------------------------------------------------------------


@REM 			SALVATAGGIO DATI *** FINE ***
@REM *************************************************************************************************************************








