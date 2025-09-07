VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_UTILITA_Msys_Frm22_IMPORTA_OGGETTI_nw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'//==============================================================================================================//
'// CLASSE: Form_UTILITA_Msys_Frm21_IMPORTA_OGGETTI
'// VERSIONE: 2025_02_12 - COMPLETA E FUNZIONANTE
'// DESCRIZIONE:
'//   Form di utilità per importare o cancellare collettivamente le tabelle collegate (@TABELLE) del mini-progetto.
'//
'// SEGNALIBRI PRINCIPALI:
'//   @01 GetTabellaPaths_LINK_TABELLE_MSYS_DF -> Elenco tabelle e percorsi MDB
'//   @02 CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct -> Collega tabelle con log dettagliato
'//   @03 CancellaTabelleCollegate_LINK_TABELLE_MSYS_DF -> Cancella tabelle collegate senza log
'//   @04 EliminaTabellaSeEsiste_LINK_TABELLE_MSYS_DF -> Elimina tabella singola
'//   @05 FileEsiste -> Controllo esistenza file
'//   @06 CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct -> Cancella tutte le tabelle con log
'//   @UTIL_01 RefreshDatabaseView -> Aggiorna riquadro oggetti di Access
'//   @EVT_Load  Form_Load -> Inizializzazione della form
'//
'// NOTE:
'//   - Commenti originali lasciati invariati
'//   - Gestione errori standardizzata
'//==============================================================================================================//


'//*********************************************************************************************************************************//
'//                                          IMPORTA O CANCELLA TABELLE LINK                          *** INIZIO ***
'//*********************************************************************************************************************************//
'// CODICE PER LA SOSTITUZIONE COLLETTIVA
'// - Permette di cancellare o importare le @TABELLE del mini-progetto.
'// - Funzioni principali, QUESTI SONO I @CODICI@COLLETTIVI:
'//         TABELLE LINK    --->     LINK_TABELLE_MSYS_DF
'//         QUERY           --->     QUERY_MSYS_DF
'//
'// Segnalibri numerati:
'//   @01 GetTabellaPaths_LINK_TABELLE_MSYS_DF -> Elenco tabelle e percorsi MDB
'//   @02 CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct -> Collega tabelle, log dettagliato
'//   @03 CancellaTabelleCollegate_LINK_TABELLE_MSYS_DF -> Cancella tabelle collegate, senza log
'//   @04 EliminaTabellaSeEsiste_LINK_TABELLE_MSYS_DF -> Elimina tabella singola
'//   @05 FileEsiste -> Controllo esistenza file
'//   @06 CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct -> Cancella tutte le tabelle con log
'//
'//*********************************************************************************************************************************//


'//**********************************************************************************************************//
'//             COMDANDI CMD, TABELLE, QUERY FORM, REPORT, MODULI E CLASSI   *** INIZIO ***
'//**********************************************************************************************************//

' ============================================================
' COMANDO: Importa link tabelle
' ============================================================
Public Sub Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click()
    On Error GoTo Err_Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click
    
    Call CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct

Exit_Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click:
    Exit Sub

Err_Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: Cancella link tabelle
' ============================================================
Public Sub Cmd_CANCELLA_LINK_TABELLE_MSYS_DF_Click()
    On Error GoTo Err_Cmd_CANCELLA_LINK_TABELLE_MSYS_DF_Click
 
    Call CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct
       
Exit_Cmd_CANCELLA_LINK_TABELLE_MSYS_DF_Click:
    Exit Sub

Err_Cmd_CANCELLA_LINK_TABELLE_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_CANCELLA_LINK_TABELLE_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: Importa QUERY
' ============================================================
Public Sub Cmd_IMPORTA_QUERY_MSYS_DF_Click()
    On Error GoTo Err_Cmd_IMPORTA_QUERY_MSYS_DF_Click
    
    Call ImportaQuery_QUERY_MSYS_DF_PFunct

Exit_Cmd_IMPORTA_QUERY_MSYS_DF_Click:
    Exit Sub

Err_Cmd_IMPORTA_QUERY_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_IMPORTA_QUERY_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: CANCELLA QUERY
' ============================================================
Public Sub Cmd_CANCELLA_QUERY_MSYS_DF_Click()
    On Error GoTo Err_Cmd_CANCELLA_QUERY_MSYS_DF_Click
    
    Call CancellaTutte_QUERY_MSYS_DF

Exit_Cmd_CANCELLA_QUERY_MSYS_DF_Click:
    Exit Sub

Err_Cmd_CANCELLA_QUERY_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_CANCELLA_QUERY_MSYS_DF_Click
End Sub



'//**********************************************************************************************************//
'//             COMDANDI CMD, TABELLE, QUERY FORM, REPORT, MODULI E CLASSI   *** INIZIO ***
'//**********************************************************************************************************//






' ============================================================
' @01 FUNZIONE GetTabellaPaths_LINK_TABELLE_MSYS_DF
' SCOPO:
'   Restituisce una collection di Array(NomeTabella, PercorsoFile).
'   - Controlla se i file esistono
'   - Esclude righe vuote
' ============================================================
Public Function GetTabellaPaths_LINK_TABELLE_MSYS_DF() As Collection
    On Error GoTo GestErrore
    
    Dim TabellaPath As New Collection
    Dim PercorsoFile As String
    
    ' === RESET campo di appoggio (se esiste in un form) ===
    sourceDBPath_s_Txt = ""
    Me.sourceDBPath_s_Txt.Requery
    
    ' === Percorso unico ===
    PercorsoFile = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb"
    
    ' === Verifica esistenza file ===
    If Not FileEsiste(PercorsoFile) Then
        MsgBox "File mancante: " & vbCrLf & PercorsoFile, vbExclamation
        Set GetTabellaPaths_LINK_TABELLE_MSYS_DF = TabellaPath
        Exit Function
    End If
    
    ' === Aggiunta nomi tabella direttamente alla Collection ===
    TabellaPath.Add Array("MSys_{@===================================================@}_Tab", PercorsoFile)
    TabellaPath.Add Array("MSys_DF01_}----------------------------------------------------@", PercorsoFile)
    TabellaPath.Add Array("MSys_DF01_COMANDI", PercorsoFile)
    
    ' === Aggiorna campo Txt Form PATH DB  ===
    sourceDBPath_s_Txt = PercorsoFile
    
    ' === Restituzione collection ===
    Set GetTabellaPaths_LINK_TABELLE_MSYS_DF = TabellaPath
    Exit Function

GestErrore:
    MsgBox "Errore in GetTabellaPaths_LINK_TABELLE_MSYS_DF: " & Err.Description, vbExclamation, "Errore"
    Set GetTabellaPaths_LINK_TABELLE_MSYS_DF = Nothing
End Function



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

    MsgBox "Operazione LINK TABELLE completata!" & vbCrLf & _
           "Tabelle collegate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log salvato in: " & LogFileName, vbInformation

CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct_Exit:
    Exit Function

CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct_Err:
    MsgBox "Errore: " & Err.Description, vbExclamation, "Errore di collegamento tabelle"
    Resume CollegaTabelle_LINK_TABELLE_MSYS_DF_PFunct_Exit
End Function



' ============================================================
' @03 FUNZIONE CancellaTabelleCollegate_LINK_TABELLE_MSYS_DF
' SCOPO:
'   Cancella tutte le tabelle collegate definite in GetTabellaPaths.
'   - Versione rapida, senza log
' ============================================================
Public Sub CancellaTabelleCollegate_LINK_TABELLE_MSYS_DF()
    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths_LINK_TABELLE_MSYS_DF()

    If TabellaPath.Count = 0 Then
        MsgBox "La collection delle tabelle è vuota. Nessuna tabella da cancellare.", vbInformation
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To TabellaPath.Count
        Dim Item As Variant
        Item = TabellaPath(i)
        
        If Len(Item(0)) = 0 Then GoTo ProssimaTabella
        EliminaTabellaSeEsiste_LINK_TABELLE_MSYS_DF Item(0)
ProssimaTabella:
    Next i

    MsgBox "Cancellazione dei collegamenti completata!", vbInformation
End Sub



' ============================================================
' @04 FUNZIONE EliminaTabellaSeEsiste_LINK_TABELLE_MSYS_DF
' SCOPO:
'   Elimina una tabella dal database corrente se esiste
' ============================================================
Public Sub EliminaTabellaSeEsiste_LINK_TABELLE_MSYS_DF(NomeTabella As Variant)
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
' @06 FUNZIONE CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct
' SCOPO:
'   Cancella tutte le tabelle collegate definite in GetTabellaPaths.
'   - Scrive log dettagliato su file
'   - Controlla validità NomeTabella
'   - Verifica l’esistenza della tabella nel DB corrente e la elimina
'
' MODIFICHE EVIDENZIATE:
'   - Dichiarate le variabili NomeTabella e PercorsoFile come String
'   - Uso di CStr() e Trim$() per convertire in stringa Item(0) e Item(1)
'   - Passata la variabile NomeTabella a EliminaTabellaSeEsiste... al posto di Item(0)
'   - Rimosso controllo percorso file (non necessario per DB corrente)
'   - Controllo NomeTabella aggiornato: Len(NomeTabella) > 0, altrimenti SKIPPED
'   - Aggiunti commenti espliciti in ogni punto dove c’è stata una modifica
' ============================================================
Public Function CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct()
    On Error GoTo CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct_Err

    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths_LINK_TABELLE_MSYS_DF()

    ' === Controllo collection vuota ===
    If TabellaPath.Count = 0 Then
        MsgBox "La collection delle tabelle è vuota. Nessuna tabella da cancellare.", vbInformation
        Exit Function
    End If

    Dim i As Long
    Dim Item As Variant        ' deve rimanere Variant perché contiene un array
    Dim NomeTabella As String  ' variabile tipizzata per Item(0) (MODIFICA)
    Dim PercorsoFile As String ' variabile tipizzata per Item(1) (MODIFICA)
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long
    Dim tdf As DAO.TableDef
    Dim TabellaEsiste As Boolean
    Dim logFile As Integer
    Dim LogFileName As String

    ' === Inizializza log ===
    LogFileName = CurrentProject.Path & "\CANC_TutteTabelleCollegate_PFunct_LOG.txt"
    logFile = FreeFile
    Open LogFileName For Output As #logFile

    Print #logFile, "Log esecuzione funzione: CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct"
    Print #logFile, "Data/Ora: " & Now()
    Print #logFile, "--------------------------------------------"

    ConteggioOk = 0
    ConteggioErr = 0

    ' === Ciclo sugli elementi della collection ===
    For i = 1 To TabellaPath.Count
        Item = TabellaPath(i) ' ogni elemento è un array con NomeTabella e PercorsoFile

        ' --- Estrazione sicura in variabili tipizzate (MODIFICA) ---
        NomeTabella = Trim$(CStr(Item(0) & ""))
        PercorsoFile = Trim$(CStr(Item(1) & ""))
        
        
        ' --- Controllo nome tabella vuoto (MODIFICA: rimosso percorso file) ---
        If Len(NomeTabella) = 0 Then
            Print #logFile, Now & " - SKIPPED: Nome tabella vuota"
            ConteggioErr = ConteggioErr + 1
            GoTo ProssimaTabella
        End If

        ' --- Verifica se la tabella esiste nel DB corrente ---
        TabellaEsiste = False
        For Each tdf In CurrentDb.TableDefs
            If tdf.Name = NomeTabella Then
                TabellaEsiste = True
                Exit For
            End If
        Next tdf
            
        ' === Aggiorna campo Txt Form PATH DB  ===
        sourceDBPath_s_Txt = PercorsoFile
            
        
        If TabellaEsiste Then
            On Error Resume Next
            ' --- Eliminazione tabella con NomeTabella tipizzato (MODIFICA) ---
            EliminaTabellaSeEsiste_LINK_TABELLE_MSYS_DF NomeTabella
            If Err.Number = 0 Then
                Print #logFile, Now & " - OK: Eliminata " & NomeTabella
                ConteggioOk = ConteggioOk + 1
            Else
                Print #logFile, Now & " - ERRORE: " & NomeTabella & " -> " & Err.Description
                ConteggioErr = ConteggioErr + 1
                Err.Clear
            End If
            On Error GoTo 0
        Else
            Print #logFile, Now & " - NON TROVATA: " & NomeTabella
            ConteggioErr = ConteggioErr + 1
        End If

ProssimaTabella:
    Next i

    ' === Riepilogo log ===
    Print #logFile, "--------------------------------------------"
    Print #logFile, "TOTALE TABELLE ELIMINATE: " & ConteggioOk
    Print #logFile, "TOTALE ERRORI/SKIPPED: " & ConteggioErr
    Close #logFile

    ' === aggiorna riquadro oggetti ===
    Call RefreshDatabaseView

    MsgBox "Cancellazione LINK TABELLE completata!" & vbCrLf & _
           "Tabelle eliminate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log salvato in: " & LogFileName, vbInformation

CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct_Exit:
    Exit Function

CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct_Err:
    MsgBox "Errore: " & Err.Description, vbExclamation, "Errore cancellazione tabelle"
    Resume CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct_Exit
End Function


'//*********************************************************************************************************************************//
'//                                          IMPORTA O CANCELLA TABELLE LINK                          *** FINE ***
'//*********************************************************************************************************************************//






'//*********************************************************************************************************************************//
'//                                MODULO UNICO: @IMPORTA@QUERY  O @CANCELLA@ QUERY                                                 //
'//*********************************************************************************************************************************//


' ============================================================
' FUNZIONE GetQueryPaths_QUERY_MSYS_DF
' SCOPO:
'   Restituisce una collection di Array(NomeQuery, PercorsoFile)
'   che definisce le query da collegare dal DB sorgente.
' ============================================================
Public Function GetQueryPaths_QUERY_MSYS_DF() As Collection
    On Error GoTo GestErrore
    
    Dim QueryPath As New Collection
    Dim PercorsoFile As String
    
    ' === Percorso del file sorgente contenente le query ===
    PercorsoFile = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb"
    
    If Not FileEsiste(PercorsoFile) Then
        MsgBox "File mancante: " & vbCrLf & PercorsoFile, vbExclamation
        Set GetQueryPaths_QUERY_MSYS_DF = QueryPath
        Exit Function
    End If
    
    ' === Elenco query da collegare ===
    QueryPath.Add Array("MSys_{@-----------------------OGGETTI_DB----------------------@}", PercorsoFile)
    QueryPath.Add Array("MSys_DF01_COMANDI_Qry01_01_SELECT_TUTTI", PercorsoFile)
    QueryPath.Add Array("Msys_DF01_COMANDI_Qry01_02_GROUP_TIPO", PercorsoFile)
    
    Set GetQueryPaths_QUERY_MSYS_DF = QueryPath
    
    Exit Function

GestErrore:
    MsgBox "Errore in GetQueryPaths_QUERY_MSYS_DF: " & Err.Description, vbExclamation, "Errore"
    Set GetQueryPaths_QUERY_MSYS_DF = Nothing
End Function


' ============================================================
' FUNZIONE ImportaQuery_QUERY_MSYS_DF_PFunct
' SCOPO:
'   Collega tutte le query definite in GetQueryPaths_QUERY_MSYS_DF
'   - Scrive log dettagliato
' ============================================================
Public Function ImportaQuery_QUERY_MSYS_DF_PFunct()
    On Error GoTo ErrHandler

    Dim QueryPath As Collection
    Set QueryPath = GetQueryPaths_QUERY_MSYS_DF()
    
    Dim NomeQuery As String
    Dim PercorsoFile As String
    Dim i As Long
    Dim logFile As Integer
    Dim LogFileName As String
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long

    If QueryPath.Count = 0 Then
        MsgBox "Nessuna query da collegare.", vbInformation
        Exit Function
    End If
    
    ' File log
    LogFileName = CurrentProject.Path & "\CollegaQuery_LOG.txt"
    logFile = FreeFile
    Open LogFileName For Output As #logFile
    Print #logFile, "Log esecuzione funzione: ImportaQuery_QUERY_MSYS_DF_PFunct"
    Print #logFile, "Data/Ora: " & Now()
    Print #logFile, "--------------------------------------------"

    ' Ciclo
    For i = 1 To QueryPath.Count
        Dim Item As Variant
        Item = QueryPath(i)
        NomeQuery = Item(0)
        PercorsoFile = Item(1)
        
        If Len(NomeQuery) = 0 Then
            Print #logFile, Now & " - SKIPPED: Nome query vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        If Not FileEsiste(PercorsoFile) Then
            Print #logFile, Now & " - ERRORE: File non trovato -> " & PercorsoFile
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        ' Elimina se già esiste
        EliminaQuerySeEsiste_QUERY_MSYS_DF NomeQuery
        
        ' Collega
        On Error Resume Next
        
        '//@cmd@importa@QUERY
        DoCmd.TransferDatabase _
            TransferType:=acImport, _
            DatabaseType:="Microsoft Access", _
            DatabaseName:=PercorsoFile, _
            ObjectType:=acQuery, _
            Source:=NomeQuery, _
            Destination:=NomeQuery, _
            StructureOnly:=False

        If Err.Number = 0 Then
            Print #logFile, Now & " - OK: " & NomeQuery
            ConteggioOk = ConteggioOk + 1
        Else
            Print #logFile, Now & " - ERRORE: " & NomeQuery & " -> " & Err.Description
            ConteggioErr = ConteggioErr + 1
            Err.Clear
        End If
        On Error GoTo 0
Prossimo:
    Next i
    
    ' Riepilogo
    Print #logFile, "--------------------------------------------"
    Print #logFile, "Totale collegate: " & ConteggioOk
    Print #logFile, "Totale errori/skipped: " & ConteggioErr
    Close #logFile
    
     ' === aggiorna riquadro oggetti ===
    Call RefreshDatabaseView

    MsgBox "Operazione IMPORTA QUERY completata!" & vbCrLf & _
           "Collegate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log: " & LogFileName, vbInformation
    Exit Function

ErrHandler:
    MsgBox "Errore in ImportaQuery_QUERY_MSYS_DF_PFunct: " & Err.Description, vbExclamation
End Function


' ============================================================
' SUB CancellaQueryCollegate_QUERY_MSYS_DF
' SCOPO:
'   Cancella tutte le query collegate definite in GetQueryPaths_QUERY_MSYS_DF
' ============================================================
Public Sub CancellaQueryCollegate_QUERY_MSYS_DF()
    Dim QueryPath As Collection
    Set QueryPath = GetQueryPaths_QUERY_MSYS_DF()
    
    If QueryPath.Count = 0 Then Exit Sub
    
    Dim i As Long
    Dim Item As Variant
    Dim NomeQuery As String
    
    For i = 1 To QueryPath.Count
        ' Recupero l'array da Collection
        Item = QueryPath(i)
        
        ' Copio l’elemento in una variabile String per evitare errore "by reference"
        NomeQuery = CStr(Item(0))
        
        If Len(NomeQuery) > 0 Then
            EliminaQuerySeEsiste_QUERY_MSYS_DF NomeQuery
        End If
    Next i
    
    MsgBox "Cancellazione QUERY completata!", vbInformation
    
    '//TODO : manca la gestione errori routine
    
End Sub


' ============================================================
' FUNZIONE CancellaTutte_QUERY_MSYS_DF
' SCOPO:
'   Cancella tutte le query collegate definite in GetQueryPaths_QUERY_MSYS_DF
'   - Scrive log dettagliato
' ============================================================
Public Function CancellaTutte_QUERY_MSYS_DF()
    On Error GoTo ErrHandler
    
    Dim QueryPath As Collection
    Set QueryPath = GetQueryPaths_QUERY_MSYS_DF()
    
    If QueryPath.Count = 0 Then
        MsgBox "Nessuna query da cancellare.", vbInformation
        Exit Function
    End If
    
    Dim i As Long, Item As Variant
    Dim NomeQuery As String
    Dim ConteggioOk As Long, ConteggioErr As Long
    Dim logFile As Integer, LogFileName As String
    
    LogFileName = CurrentProject.Path & "\CancellaQuery_LOG.txt"
    logFile = FreeFile
    Open LogFileName For Output As #logFile
    
    Print #logFile, "Log esecuzione funzione: CancellaTutte_QUERY_MSYS_DF"
    Print #logFile, "Data/Ora: " & Now()
    Print #logFile, "--------------------------------------------"
    
    For i = 1 To QueryPath.Count
        Item = QueryPath(i)
        NomeQuery = Trim$(CStr(Item(0)))
        
        If Len(NomeQuery) = 0 Then
            Print #logFile, Now & " - SKIPPED: Nome vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        If QueryEsiste_QUERY_MSYS_DF(NomeQuery) Then
            EliminaQuerySeEsiste_QUERY_MSYS_DF NomeQuery
            If Not QueryEsiste_QUERY_MSYS_DF(NomeQuery) Then
                Print #logFile, Now & " - OK: Eliminata " & NomeQuery
                ConteggioOk = ConteggioOk + 1
            Else
                Print #logFile, Now & " - ERRORE: Non eliminata " & NomeQuery
                ConteggioErr = ConteggioErr + 1
            End If
        Else
            Print #logFile, Now & " - NON TROVATA: " & NomeQuery
            ConteggioErr = ConteggioErr + 1
        End If
Prossimo:
    Next i
    
    Print #logFile, "--------------------------------------------"
    Print #logFile, "Totale eliminate: " & ConteggioOk
    Print #logFile, "Totale errori/skipped: " & ConteggioErr
    Close #logFile
    
    Call RefreshDatabaseView
    MsgBox "Cancellazione completata!" & vbCrLf & _
           "Eliminate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log: " & LogFileName, vbInformation
    Exit Function
    
ErrHandler:
    MsgBox "Errore in CancellaTutte_QUERY_MSYS_DF: " & Err.Description, vbExclamation
End Function


' ============================================================
' SUB EliminaQuerySeEsiste_QUERY_MSYS_DF
' SCOPO:
'   Elimina una query dal database corrente se esiste
' ============================================================
Public Sub EliminaQuerySeEsiste_QUERY_MSYS_DF(NomeQuery As String)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete NomeQuery
    On Error GoTo 0
End Sub


' ============================================================
' FUNZIONE QueryEsiste_QUERY_MSYS_DF
' SCOPO:
'   Verifica se una query esiste nel DB corrente
' ============================================================
Public Function QueryEsiste_QUERY_MSYS_DF(NomeQuery As String) As Boolean
    Dim qdf As DAO.QueryDef
    QueryEsiste_QUERY_MSYS_DF = False
    For Each qdf In CurrentDb.QueryDefs
        If qdf.Name = NomeQuery Then
            QueryEsiste_QUERY_MSYS_DF = True
            Exit For
        End If
    Next qdf
End Function




'//*********************************************************************************************************************************//
'//                                MODULO UNICO: @IMPORTA@QUERY  O @CANCELLA@ QUERY  *** FINE ***                                   //
'//*********************************************************************************************************************************//







'//*********************************************************************************************************************************//
'//                                          @FUNZIONI@COMUNI A TUTTO IL PROGETTO                     *** INIZIO ***
'//*********************************************************************************************************************************//

' ============================================================
' @UTIL_01 RefreshDatabaseView
' SCOPO:
'   Aggiorna il riquadro oggetti di Access (Navigation Pane)
'
' Metodo consigliato: Application.RefreshDatabaseWindow
' - Forza l’aggiornamento del riquadro oggetti (tabelle, query, form, report, macro, moduli)
' - Utile dopo operazioni di importazione, eliminazione o modifica oggetti
'
' Note:
' - Alcune versioni recenti di Access eseguono refresh automatico, ma con ritardo.
' - In alternativa si può usare anche DoCmd.RunCommand acCmdRefresh come fallback.
' ============================================================
Public Sub RefreshDatabaseView()
    On Error Resume Next
    
    ' === Metodo principale: refresh Navigation Pane ===
    Application.RefreshDatabaseWindow
    
    ' === Metodo alternativo (fallback) ===
    'DoCmd.RunCommand acCmdRefresh
    
    On Error GoTo 0
End Sub


' ============================================================
' @05 FUNZIONE FileEsiste
' SCOPO:
'   Funzione di utilità che controlla se un file esiste.
'   Restituisce TRUE se esiste, FALSE altrimenti.
' ============================================================
Public Function FileEsiste(PercorsoFile As String) As Boolean
    On Error Resume Next
    FileEsiste = (Dir(PercorsoFile, vbNormal) <> "")
    On Error GoTo 0
End Function




'//*********************************************************************************************************************************//
'//                                          @FUNZIONI@COMUNI A TUTTO IL PROGETTO                     *** FINE ***
'//*********************************************************************************************************************************//






'//*********************************************************************************************************************************//
'//                                          @EVENTI@FORM A TUTTO IL PROGETTO                     *** INIZIO ***
'//*********************************************************************************************************************************//


' =======================================================================================
' Evento: Form_Load
' Descrizione: viene eseguito automaticamente all'apertura della form.
'              Qui puoi inizializzare variabili, impostare proprietà,
'              aprire connessioni o caricare dati di avvio.
' =======================================================================================
Private Sub Form_Load()
    On Error GoTo ErrHandler   ' Attiva la gestione degli errori

    ' --------------------------------------------------------------
    ' Inizializzazione delle variabili e impostazioni della form
    ' --------------------------------------------------------------
    Me.Caption = "Gestione Oggetti"   ' Imposta il titolo della finestra
    ' TODO: aggiungi qui altre inizializzazioni necessarie
    ' ad esempio: caricamento combo, settaggio default, ecc.

' --------------------------------------------------------------
' uscita + Gestione errori
' --------------------------------------------------------------


Exit Sub   ' Uscita normale dal metodo

ErrHandler:
    MsgBox "Errore durante il caricamento della form: " & Err.Description, _
           vbCritical, "Errore " & Err.Number
    Resume Next   ' Continua l'esecuzione dopo l'errore (valuta Resume/Exit in base al caso)
End Sub








'//*********************************************************************************************************************************//
'//                                          @EVENTI@FORM A TUTTO IL PROGETTO                     *** FINE ***
'//*********************************************************************************************************************************//


