@REM FILE ----> BAT_RICREA_PATH_MENU.BAT
            @echo off
            setlocal enableextensions

            rem Percorso completo da ricostruire in MSys_DF03_02 al posto dell'originale MSys_DF03
            set "TARGET=c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF03_02\"

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
