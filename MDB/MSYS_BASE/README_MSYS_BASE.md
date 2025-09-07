# MSYS_BASE
   progetto base per i mini progetti








CHATGPT

   SEGNALIBRI
      TI POSSO DARE LA CLASSE della form gestione oggetti per un 
      controllo generale di implementazione tra le funzione ed 
      un AGGIORNAMENTO DEI SEGNALIBRI in TESTATA 
      e sulle singole funzioni LASCIANDO I COMMENTI ORIGINALI?


   ADATTA FUNZIONE DA IMPORTA TABELLE LINK AD IMPORTA QUERY 


      ' ============================================================
' @02 FUNZIONE CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct
' SCOPO:
'   Collega tutte le tabelle definite in GetTabellaPaths.
'   - Controlla NomeTabella e PercorsoFile
'   - Scrive un log dettagliato
' ============================================================
Public Function CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct()
    On Error GoTo CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct_Err

    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths_LINK_TABELLE_MSYS_DF()
    
    Dim NomeTabella As String
    Dim PercorsoFile As String
    Dim i As Long
    Dim logFile As Integer
    Dim LogFileName As String
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long

    If TabellaPath.Count = 0 Then
        MsgBox "La collection delle tabelle è vuota. Nessuna operazione effettuata.", vbInformation
        Exit Function
    End If

    ' Creazione file log
    LogFileName = CurrentProject.Path & "\CollegaTabelle_LINK_TABELLE_MSYS_DF_LOG.txt"
    logFile = FreeFile
    Open LogFileName For Output As #logFile

    Print #logFile, "Log esecuzione funzione: CollegaTabelle_LINK_TABELLE_MSYS_DF"
    Print #logFile, "Data/Ora: " & Now()
    Print #logFile, "--------------------------------------------"

    ConteggioOk = 0
    ConteggioErr = 0

    ' Ciclo su tutte le tabelle
    For i = 1 To TabellaPath.Count
        Dim Item As Variant
        Item = TabellaPath(i)
        NomeTabella = Item(0)
        PercorsoFile = Item(1)

        ' Controllo valori vuoti
        If Len(NomeTabella) = 0 Or Len(PercorsoFile) = 0 Then
            Print #logFile, Now & " - SKIPPED: Nome tabella o percorso vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo ProssimaTabella
        End If

        ' Controllo percorso file
        If Not FileEsiste(PercorsoFile) Then
            Print #logFile, Now & " - ERRORE: Percorso file NON valido per " & NomeTabella & " -> " & PercorsoFile
            ConteggioErr = ConteggioErr + 1
            GoTo ProssimaTabella
        End If

        ' Elimina tabella se già esiste, poi crea collegamento
        On Error Resume Next
        EliminaTabellaSeEsiste_LINK_TABELLE_MSYS_DF NomeTabella
        DoCmd.TransferDatabase acLink, "Microsoft Access", PercorsoFile, acTable, NomeTabella, NomeTabella, False

        If Err.Number = 0 Then
            Print #logFile, Now & " - OK: " & NomeTabella & " -> " & PercorsoFile
            ConteggioOk = ConteggioOk + 1
        Else
            Print #logFile, Now & " - ERRORE: " & NomeTabella & " -> " & Err.Description
            ConteggioErr = ConteggioErr + 1
            Err.Clear
        End If
ProssimaTabella:
        On Error GoTo 0
    Next i

    ' Riepilogo log
    Print #logFile, "--------------------------------------------"
    Print #logFile, "TOTALE TABELLE COLLEGATE: " & ConteggioOk
    Print #logFile, "TOTALE ERRORI/SKIPPED: " & ConteggioErr
    Close #logFile

    MsgBox "Operazione completata!" & vbCrLf & _
           "Tabelle collegate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log salvato in: " & LogFileName, vbInformation

CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct_Exit:
    Exit Function

CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct_Err:
    MsgBox "Errore: " & Err.Description, vbExclamation, "Errore di collegamento tabelle"
    Resume CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct_Exit
End Function

