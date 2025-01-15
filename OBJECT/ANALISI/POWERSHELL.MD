POWERSHELL.MD

	NOTE
		i comandi di powershelle da sistermare in archivio tutorial, per orma viene salvato
		in questa directory, poi è da posizionare in modo completo e definitivo.

			@powershell_(@shelle di @windows in alternativa al dos)


	COMANDI
		A
			APERTURA DELL'IDE DI POWERSHELL
				powershell ISE, i comando che apre l'ide di power shell.
				Il comando funzione sia in DOS che nello script di Powershell



		P
			powershell ISE
				apro l'ide di powershell

		E
			@EOLO
				Per le proprieta della rete utilizzare :
   					SICUREZZA
   						Tipo di sicurezza 	: WPA2-Personal
   						Tipo di crittografia : AES
     					Chiave di sicurezza	: en8PCSGkb

					PROBLEMI
						se Eolo non è visibile nel centro di 
						connessione alla rete avviare il controllo di eolo della rete sul router
						chimato AUTODIAGNOSI





		C
			CHIUDI_PROCESSI_EDGE_PADRE+FIGLI


				Per gestire sia il processo padre che i processi figli in modo corretto, possiamo modificare lo script in modo che chiuda tutti i processi correlati a un dato processo padre.

				Ecco uno script aggiornato che chiude il processo padre e tutti i relativi processi figli:

				Script aggiornato per chiudere processi padre e figli
				powershell
				
				Modifiche apportate:
				Get-WmiObject: Utilizzo di Win32_Process per trovare i processi figli. Questo comando controlla la proprietà ParentProcessId per trovare i processi che sono figli del processo padre.
				Chiudi prima i processi figli: Nella funzione ChiudiProcessiEdgeConFigli, i processi figli vengono chiusi prima di chiudere il processo padre.
				Gestione completa di padre e figli: Il ciclo prima chiude i processi figli e successivamente termina il processo padre.
				Come funziona lo script:
				Viene eseguita una ricerca di tutti i processi msedge.exe (sia padri che figli).
				Per ogni processo, si cercano i processi figli associati.
				Prima si chiudono i processi figli, poi il processo padre.
				Esecuzione dello script:
				Copia il codice.
				Apri PowerShell come amministratore.
				Incolla e esegui il codice oppure salvalo come .ps1 e avvialo.
				PowerShell chiuderà tutti i processi msedge.exe, inclusi quelli padre e figli.


				# Funzione per cercare e contare i processi msedge.exe
				function TrovaProcessiEdge {
				    $processiEdge = Get-Process -Name "msedge" -ErrorAction SilentlyContinue

				    if ($processiEdge) {
				        $numeroProcessi = $processiEdge.Count
				        Write-Host "Trovati $numeroProcessi processi msedge.exe in esecuzione."
				        return $processiEdge
				    } else {
				        Write-Host "Nessun processo msedge.exe trovato."
				        return $null
				    }
				}

				# Funzione per chiudere un processo e i suoi figli
				function ChiudiProcessiEdgeConFigli {
				    param (
				        [array]$processiDaChiudere
				    )

				    foreach ($processo in $processiDaChiudere) {
				        try {
				            # Ottieni i processi figli del processo corrente
				            $figli = Get-WmiObject Win32_Process | Where-Object { $_.ParentProcessId -eq $processo.Id }

				            # Chiudi prima i processi figli
				            foreach ($figlio in $figli) {
				                Stop-Process -Id $figlio.ProcessId -Force
				                Write-Host "Processo figlio con ID $($figlio.ProcessId) terminato."
				            }

				            # Ora chiudi il processo padre
				            Stop-Process -Id $processo.Id -Force
				            Write-Host "Processo padre con ID $($processo.Id) terminato."
				        } catch {
				            Write-Host "Errore durante la chiusura del processo con ID $($processo.Id): $_"
				        }
				    }
				}

				# Esecuzione dello script principale
				$processiEdge = TrovaProcessiEdge

				if ($processiEdge) {
				    ChiudiProcessiEdgeConFigli -processiDaChiudere $processiEdge
				} else {
				    Write-Host "Non ci sono processi msedge.exe da chiudere."
				}










			ESECUZIONE_DOS
				@echo off
					powershell.exe -ExecutionPolicy Bypass -File "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\POWERSHELL\CHIUDI_EDGE_ProcessiPadre+Figli.ps1"
					pause

			ESECUZIONE_POWERSHELL
				powershell.exe -ExecutionPolicy Bypass -File "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\POWERSHELL\CHIUDI_EDGE_ProcessiPadre+Figli.ps1"



		W
			WIFI RESET SCHEDA
				reimposta la scheda di rete
				in cmd impostare la directory
				cd c:\Windows\sistem32\ 
				comando	
					netsh int ip reset resetlog.txt



