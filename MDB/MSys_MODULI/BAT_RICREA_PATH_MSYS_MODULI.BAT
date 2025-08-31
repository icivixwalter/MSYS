@REM FILE ----> BAT_RICREA_PATH_MENU.BAT
            @echo off
            setlocal enableextensions

            rem Percorso completo da ricostruire
            set "TARGET=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\"

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
