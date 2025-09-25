@REM FILE ----> BAT_RICREA_PATH_MENU.BAT
            @echo off
            setlocal enableextensions

            rem Percorso completo da ricostruire DA MSys_DF10 A MSys_DF10_2 PER EVITARE in caso
            rem di attivazione sovrascitturazioni.
            rem Percorso completo puo essere effettuato solo in caso di inesistenza della directory
            rem principale denominata MSys_DF10.
            set "TARGET=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF10_2\"

            if exist "%TARGET%" (
                echo CARTELLA GIA ESISTENTE: "%TARGET%"
                timeout /t 5 /nobreak


            ) else (
                echo Creo: "%TARGET%"
                mkdir "%TARGET%"
                if errorlevel 1 (
                    echo ERRORE: impossibile creare "%TARGET%".
                    exit /b 1
                ) else (
                    echo Creato con successo.
                     timeout /t 5 /nobreak >nul
                )
            )

            endlocal
