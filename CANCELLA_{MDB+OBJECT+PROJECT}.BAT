@REM CANCELLA MDB OBJEC E PROJECT
@REM .
@REM . PARAMETRI DA SOSTITUIRE 
@REM . FILE ----->  Project_MSYS
@REM . PATH ----->  c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\
@REM .


ECHO OFF

@REM  cd.. devi posizionarti su git_desktop
@REM imposto la variabile path:
SET PATH_S=c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\


@CD %PATH_S%
@dir /W

@ECHO  CANCELLO I DATI
@ECHO  .......................................................
@ECHO  .
@ECHO  .
@ECHO  .

@DIR MDB


::           ELIMINA LE CARTELLE  CON I PARAMETRI
::========================================================================================
:: Elimina le cartelle specificate (se esistono)
:: /S elimina tutte le sottocartelle e file contenuti
:: /Q evita la richiesta di conferma
:: rd /S /Q "[MDB]"
:: rd /S /Q "[OBJECT]"

@echo 				ATTIVO LA ELIMINAZIONE DELLE CARTELLE  *** INIZIO ***
@echo --------------------------------------------------------------------
@echo 
@echo 
@echo ELIMINO LA CARTELLA MDB E OBJECT

 
	@REM  elimina una directory (RMDIR) + le sottodirectory (/S) senza conferma (/Q)
	rmdir /s /q "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB"



@DIR OBJECT

	@REM  elimina una directory (RMDIR) + le sottodirectory (/S) senza conferma (/Q)
	rmdir /s /q "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT"

DIR


:: Aggiorna la schermata di Windows Explorer
:: Il seguente comando forza un refresh della shell di Windows
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"


@echo 				  *** FINE ***
@echo --------------------------------------------------------------------

@ECHO  pausa  per consentire il commit
@ECHO
@TIMEOUT /T 9 /NOBREAK



::           ELIMINA I FILE CON I PARAMETRI
::========================================================================================
:: Elimina i file specificati
:: /F forza la cancellazione dei file di sola lettura
:: /Q evita la richiesta di conferma

@echo --------------------------------------------------------------------
@echo 				ATTIVO LA ELIMINAZIONE DEI FILE  *** INIZIO ***
@echo 
@echo 
@echo 

del "AA__________________________________________________________________________AA.txt" /F /Q
del "git_OpenGitHub.bat" /F /Q
del "git_OPEN_OVERLEAF.BAT" /F /Q
del "git_PULL.BAT" /F /Q
del "git_PULL_LOG.txt" /F /Q
del "git_STATUS+ADD+COMMIT+PUSH.BAT" /F /Q
del "git_STATUS.BAT" /F /Q
del "git_STATUS_LOG.txt" /F /Q
del "Project_MSYS.sublime-project" /F /Q
del "Project_MSYS.sublime-workspace" /F /Q
del "README.md" /F /Q
del "ZIP_7ZIP_INCREMENTALE_{CRIPTA+DECODIFICA_INTERNO}.BAT" /F /Q
del "ZIP_7ZIP_________________________________________________________________ZZ.txt" /F /Q
del "Zip_N70_CRIPTA.bat" /F /Q
del "Zip_N70_DECODIFICA.bat" /F /Q
del "Zip_N70_ElencoFile.txt" /F /Q
del "Zip_N70_ElencoFileESCLUSI.txt" /F /Q
del "Zip_N70_RAR_____________________________________________________________RAR.txt" /F /Q
del "Zip_N70____________________________________________CRIPTA_DECODIFICA.txt" /F /Q
del "ZIP_RAR.BAT" /F /Q
del "Zip_rar_C.bat" /F /Q
del "Zip_rar_C_GitHub.bat" /F /Q
del "ZZ__________________________________________________________________________ZZ.txt" /F /Q

@echo 
@echo 
@echo 			Operazione completata.   *** FINE ***
@echo --------------------------------------------------------------------

:: Il seguente comando forza un refresh della shell di Windows
:: Aggiorna la schermata di Windows Explorer
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"

@ECHO  pausa  per consentire il commit
@ECHO
@TIMEOUT /T 5 /NOBREAK





@echo 		ALLINEO LA REPOSITORY LOCALE CON GITHUB  *** INIZIO ***
@echo --------------------------------------------------------------------
@echo 
@echo 
@echo 

	@REM Aggiorna le informazioni dal remoto
	git fetch origin	
	
	@REM Sovrascrive completamente i file locali con quella dalla branch remota
	git reset --hard origin/calcolo

	@REM Verifica che la directory locale sia ora identica al repository remoto 
	git status
	

@ECHO  .
@ECHO  .        *** FINE ***
@ECHO  .......................................................


:: Aggiorna la schermata di Windows Explorer
:: Il seguente comando forza un refresh della shell di Windows
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"
powershell.exe -command "(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() }"



@REM 01)
@REM sospensione con crononometro 
@REM ---->	timeout /t 2 /nobreak > NUL
@REM ---->
@ECHO
@ECHO
@ECHO  pausa  per consentire il commit
@ECHO
@TIMEOUT /T 5 /NOBREAK

