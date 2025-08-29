
'//*************************************************************************************************//'
'//*
'//* MODULO : Msys_DF01_Mdl01_COLLEGA_COMANDI
'//*
'//* Per ricapitolare, il modulo rimane così strutturato:
'//* @01 GetTabellaPaths ? contiene solo i dati (nomi tabelle e percorsi MDB).
'//* @02 CollegaTabelle_PFunct ? collega le tabelle, crea log e messaggi.
'//* @03 CancellaTabelleCollegate ? elimina i collegamenti senza log né conferma.
'//* @04 EliminaTabellaSeEsiste ? verifica ed elimina una tabella se esiste.
'//* @05 CancellaTutteTabelleCollegate_PFunct ? versione avanzata con log dettagliato.
'//*
'//* Variazioni apportate:
'//* - Controllo se la collection è vuota e uscita anticipata.
'//* - Controllo se NomeTabella o PercorsoFile sono vuoti e segnalazione SKIPPED.
'//* - Aggiornamento messaggi e log per evidenziare elementi saltati.
'//*
'//*************************************************************************************************//'

' ============================================================
' @01 FUNZIONE GetTabellaPaths
' Restituisce la collection di tabelle e percorsi MDB da collegare
' Non vengono salvate nella collection le riga nulle.
' ============================================================
Public Function GetTabellaPaths() As Collection
    On Error GoTo GestErrore
    Dim TabellaPath As New Collection
    Dim NomeTabella As String
    Dim PercorsoFile As String

    ' === Primo gruppo ===
    NomeTabella = "MSys_DF01_}----------------------------------------------------@"
    PercorsoFile = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb"
    If NomeTabella <> "" And PercorsoFile <> "" Then TabellaPath.Add Array(NomeTabella, PercorsoFile)

    NomeTabella = "MSys_DF01_COMANDI"
    PercorsoFile = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb"
    If NomeTabella <> "" And PercorsoFile <> "" Then TabellaPath.Add Array(NomeTabella, PercorsoFile)

    ' === Secondo gruppo ===
  
    ' === Restituzione della collection ===
    Set GetTabellaPaths = TabellaPath
    Exit Function

GestErrore:
    MsgBox "Errore in GetTabellaPaths: " & Err.Description, vbExclamation, "Errore"
    Set GetTabellaPaths = Nothing
End Function


' ============================================================
' @02 FUNZIONE CollegaTabelle_PFunct
' Collega le tabelle del database corrente ai MDB esterni
' ============================================================
Public Function CollegaTabelle_PFunct()

    On Error GoTo CollegaTabelle_PFunct_Err

    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths()
    
    Dim NomeTabella As String
    Dim PercorsoFile As String
    Dim i As Long
    Dim LogFile As Integer
    Dim LogFileName As String
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long

    ' ============================
    ' Variazione: controllo collection vuota
    ' ============================
    If TabellaPath.Count = 0 Then
        MsgBox "La collection delle tabelle è vuota. Nessuna operazione effettuata.", vbInformation
        Exit Function
    End If

    ' @02.01 Creazione file log
    LogFileName = CurrentProject.Path & "\Msys_DF01_Mdl01_COLLEGA_COMANDI_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile

    Print #LogFile, "Log esecuzione funzione: Msys_DF01_Mdl01_COLLEGA_COMANDI"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ConteggioOk = 0
    ConteggioErr = 0

    ' @02.02 Ciclo su tutte le tabelle
    For i = 1 To TabellaPath.Count
        Dim Item As Variant
        Item = TabellaPath(i)            ' Array contenente NomeTabella e PercorsoFile
        NomeTabella = Item(0)
        PercorsoFile = Item(1)

        ' ============================
        ' Variazione: controllo valori vuoti
        ' ============================
        If Len(NomeTabella) = 0 Or Len(PercorsoFile) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome tabella o percorso vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo ProssimaTabella
        End If

        On Error Resume Next
        EliminaTabellaSeEsiste NomeTabella
        DoCmd.TransferDatabase acLink, "Microsoft Access", PercorsoFile, acTable, NomeTabella, NomeTabella, False

        If Err.Number = 0 Then
            Print #LogFile, Now & " - OK: " & NomeTabella & " -> " & PercorsoFile
            ConteggioOk = ConteggioOk + 1
        Else
            Print #LogFile, Now & " - ERRORE: " & NomeTabella & " -> " & Err.Description
            ConteggioErr = ConteggioErr + 1
            Err.Clear
        End If
ProssimaTabella:
        On Error GoTo 0
    Next i

    ' @02.03 Riepilogo log
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "TOTALE TABELLE COLLEGATE: " & ConteggioOk
    Print #LogFile, "TOTALE TABELLE NON COLLEGATE/ERRATE: " & ConteggioErr
    Close #LogFile

    ' @02.04 Messaggio a video
    MsgBox "Operazione completata!" & vbCrLf & _
           "Tabelle collegate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log salvato in: " & LogFileName, vbInformation

CollegaTabelle_PFunct_Exit:
    Exit Function

CollegaTabelle_PFunct_Err:
    MsgBox "Errore: " & Err.Description, vbExclamation, "Errore di collegamento tabelle"
    Resume CollegaTabelle_PFunct_Exit
End Function


' ============================================================
' @03 FUNZIONE CancellaTabelleCollegate
' Elimina i collegamenti alle tabelle presenti nel database corrente
' ============================================================
Public Sub CancellaTabelleCollegate()
    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths()

    ' ============================
    ' Variazione: controlla collection vuota
    ' ============================
    If TabellaPath.Count = 0 Then
        MsgBox "La collection delle tabelle è vuota. Nessuna tabella da cancellare.", vbInformation
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To TabellaPath.Count
        Dim Item As Variant
        Item = TabellaPath(i)
        
        ' Variazione: salto nomi tabella vuoti
        If Len(Item(0)) = 0 Then GoTo ProssimaTabella

        EliminaTabellaSeEsiste Item(0)
ProssimaTabella:
    Next i

    MsgBox "Cancellazione dei collegamenti completata!", vbInformation
End Sub


' ============================================================
' @04 FUNZIONE EliminaTabellaSeEsiste
' Controlla se la tabella esiste e la elimina
' ============================================================
Public Sub EliminaTabellaSeEsiste(NomeTabella As Variant)
    On Error Resume Next
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = NomeTabella Then
            CurrentDb.TableDefs.Delete NomeTabella
            Exit For
        End If
    Next tdf
    On Error GoTo 0
End Sub


' ============================================================
' @05 FUNZIONE CancellaTutteTabelleCollegate_PFunct
' Cancella tutte le tabelle collegate definite in GetTabellaPaths
' solo se effettivamente esistono nel db corrente
' con log dei risultati su file
' ============================================================
Function CancellaTutteTabelleCollegate_PFunct()
    On Error GoTo CancellaTutteTabelleCollegate_PFunct_Err

    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths()

    ' ============================
    ' Variazione: controllo collection vuota
    ' ============================
    If TabellaPath.Count = 0 Then
        MsgBox "La collection delle tabelle è vuota. Nessuna tabella da cancellare.", vbInformation
        Exit Function
    End If

    Dim i As Long
    Dim Item As Variant
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long
    Dim tdf As DAO.TableDef
    Dim TabellaEsiste As Boolean
    Dim LogFile As Integer
    Dim LogFileName As String

    ' Inizializza log
    LogFileName = CurrentProject.Path & "\CANC_TutteTabelleCollegate_PFunct_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile

    Print #LogFile, "Log esecuzione funzione: CancellaTutteTabelleCollegate_PFunct"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ConteggioOk = 0
    ConteggioErr = 0

    ' Ciclo per eliminare tutte le tabelle collegate
    For i = 1 To TabellaPath.Count
        Item = TabellaPath(i)

        ' Variazione: salto nomi tabella vuoti
        If Len(Item(0)) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome tabella vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo ProssimaTabella
        End If

        TabellaEsiste = False
        For Each tdf In CurrentDb.TableDefs
            If tdf.Name = Item(0) Then
                TabellaEsiste = True
                Exit For
            End If
        Next tdf

        If TabellaEsiste Then
            On Error Resume Next
            EliminaTabellaSeEsiste Item(0)
            If Err.Number = 0 Then
                Print #LogFile, Now & " - OK: " & Item(0)
                ConteggioOk = ConteggioOk + 1
            Else
                Print #LogFile, Now & " - ERRORE: " & Item(0) & " -> " & Err.Description
                ConteggioErr = ConteggioErr + 1
                Err.Clear
            End If
            On Error GoTo 0
        Else
            Print #LogFile, Now & " - NON TROVATA: " & Item(0)
            ConteggioErr = ConteggioErr + 1
        End If
ProssimaTabella:
    Next i

    ' Riepilogo log
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "TOTALE TABELLE ELIMINATE: " & ConteggioOk
    Print #LogFile, "TOTALE TABELLE NON TROVATE/SKIPPED: " & ConteggioErr
    Close #LogFile

    MsgBox "Cancellazione completata!" & vbCrLf & _
           "Tabelle eliminate: " & ConteggioOk & vbCrLf & _
           "Tabelle non trovate/skipped: " & ConteggioErr & vbCrLf & _
           "Log salvato in: " & LogFileName, vbInformation

CancellaTutteTabelleCollegate_PFunct_Exit:
    Exit Function

CancellaTutteTabelleCollegate_PFunct_Err:
    MsgBox "Errore: " & Err.Description, vbExclamation, "Errore di collegamento tabelle"
    Resume CancellaTutteTabelleCollegate_PFunct_Exit
End Function


