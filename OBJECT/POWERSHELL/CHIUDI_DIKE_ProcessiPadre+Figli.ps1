

# START "ATTIVA POWERSHELL" call %windir%\system32\WindowsPowerShell\v1.0\PowerShell_ISE.exe

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
