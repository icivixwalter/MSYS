@REM FILE ----> BAT_RICREA_PATH_MSYS.BAT
            @echo off
            setlocal enableextensions

            rem Percorso completo da ricostruire
            set "TARGET=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS"

            if exist "%TARGET%" (
                echo Esiste già: "%TARGET%"
            ) else (
                echo Creo: "%TARGET%"
                mkdir "%TARGET%"
                if errorlevel 1 (
                    echo ERRORE: impossibile creare "%TARGET%".
                    exit /b 1
                ) else (
                    echo Creato con successo.
                )
            )

            endlocal
