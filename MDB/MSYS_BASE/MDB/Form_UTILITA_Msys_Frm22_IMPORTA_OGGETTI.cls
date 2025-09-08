VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_UTILITA_Msys_Frm22_IMPORTA_OGGETTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//DENOMINAZIONE    -> Modello_Funzione_Pfunct
Option Compare Database
'//Variabili di database e GENERALI
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'DAO
Dim DaoDB As DAO.Database
Dim DaoWks As DAO.Workspace
Dim DaoRs As DAO.Recordset
Dim DaoRs_Lett As DAO.Recordset

'ADO
Dim ADODB As Database
Dim AdodaoRs As Recordset

'Contatori
Dim iCount As Integer
Dim iTOTcount As Integer
Dim dbl_count As Double

'Le variabili generiche
Dim sSql As String                                          '//Stringa di estrazione SQL
Dim sSql_Lett As String                                     '//Stringa di estrazione SQL

'Variabili generali
Dim Str1 As String
Dim Int1 As Integer
Dim Int2 As Integer
Dim Int3 As Integer
Dim Lng1 As Long
Dim Dbl1 As Double
Dim Bln1 As Boolean
Dim Vv1 As Variant
Dim obj1 As Object

'Gestione parametri comandi
Dim sxParamCmd_3 As String


'Larghezza e numero di colonna
Dim sLarg_Col As String
Dim iNum_Col As Integer


'//ERRORI PROCEDURA_FUNCTION O ROUTINE
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String


'//==============================================================================================================//
'// CLASSE: Form_UTILITA_Msys_Frm21_IMPORTA_OGGETTI
'// VERSIONE: 2025_02_12 - COMPLETA E FUNZIONANTE
'// DESCRIZIONE:
'//   Form di utilità per importare o cancellare collettivamente le tabelle collegate (@TABELLE) del mini-progetto.
'//
'// SEGNALIBRI PRINCIPALI:
'//   @01 GetTabellaPaths_LINK_TABELLE_MSYS_DF -> Elenco tabelle e percorsi MDB
'//   @02 Importa_LINK_TABELLE_MSYS_DF_PFunct -> Collega tabelle con log dettagliato
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
'//         TABELLE LINK        --->     LINK_TABELLE_MSYS_DF
'//         QUERY               --->     QUERY_MSYS_DF
'//         REPORT              --->     REPORT_MSYS_DF
'//         MODULI              --->     MODULI_MSYS_DF
'//         FORM                --->     FORM_MSYS_DF
'//
'// Segnalibri numerati:
'//   @01 GetTabellaPaths_LINK_TABELLE_MSYS_DF -> Elenco tabelle e percorsi MDB
'//   @02 Importa_LINK_TABELLE_MSYS_DF_PFunct -> Collega tabelle, log dettagliato
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
' COMANDO: @CMD@IMPORTA@TUTTO
' ============================================================
Public Sub Cmd_IMPORTA_TUTTI_GLI_OGGETTI_MSYS_DF_Click()
    On Error GoTo Err_Cmd_IMPORTA_TUTTI_GLI_OGGETTI_MSYS_DF_Click
    
       
        Importa_LINK_TABELLE_MSYS_DF_PFunct
        Importa_MODULI_MSYS_DF_PFunct
        Importa_QUERY_MSYS_DF_PFunct
        Importa_REPORT_MSYS_DF_PFunct
        Importa_FORM_MSYS_DF_PFunct

Exit_Cmd_IMPORTA_TUTTI_GLI_OGGETTI_MSYS_DF_Click:
    Exit Sub

Err_Cmd_IMPORTA_TUTTI_GLI_OGGETTI_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_IMPORTA_TUTTI_GLI_OGGETTI_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: @CMD@CANCELLA@TUTTO
' ============================================================
Public Sub Cmd_CANCELLA_TUTTI_GLI_OGGETTI_MSYS_DF_Click()
    On Error GoTo Err_Cmd_CANCELLA_TUTTI_GLI_OGGETTI_MSYS_DF_Click
    
    Call Importa_LINK_TABELLE_MSYS_DF_PFunct

Exit_Cmd_CANCELLA_TUTTI_GLI_OGGETTI_MSYS_DF_Click:
    Exit Sub

Err_Cmd_CANCELLA_TUTTI_GLI_OGGETTI_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_CANCELLA_TUTTI_GLI_OGGETTI_MSYS_DF_Click
End Sub




' ============================================================
' COMANDO: @CMD@CANCELLA@TABELLE@LINK
' ============================================================
Public Sub Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click()
    On Error GoTo Err_Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click
    
    Call Importa_LINK_TABELLE_MSYS_DF_PFunct

Exit_Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click:
    Exit Sub

Err_Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_IMPORTA_LINK_TABELLE_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: @CMD@CANCELLA@TABELLE@LINK
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
' COMANDO: Importa @CMD QUERY
' ============================================================
Public Sub Cmd_IMPORTA_QUERY_MSYS_DF_Click()
    On Error GoTo Err_Cmd_IMPORTA_QUERY_MSYS_DF_Click
    
    Call Importa_QUERY_MSYS_DF_PFunct

Exit_Cmd_IMPORTA_QUERY_MSYS_DF_Click:
    Exit Sub

Err_Cmd_IMPORTA_QUERY_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_IMPORTA_QUERY_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: CANCELLA @CMD QUERY
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




' ============================================================
' COMANDO: @CMD@FORM@Importa
' ============================================================
Public Sub Cmd_IMPORTA_FORM_MSYS_DF_Click()
    On Error GoTo Err_Cmd_IMPORTA_FORM_MSYS_DF_Click
    
    Call Importa_FORM_MSYS_DF_PFunct

Exit_Cmd_IMPORTA_FORM_MSYS_DF_Click:
    Exit Sub

Err_Cmd_IMPORTA_FORM_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_IMPORTA_FORM_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: @CMD@FORM@CANCELLA
' ============================================================
Public Sub Cmd_Cancella_FORM_MSYS_DF_Click()
    On Error GoTo Err_Cmd_Cancella_FORM_MSYS_DF_Click
    
    Call CancellaTutte_FORM_MSYS_DF

Exit_Cmd_Cancella_FORM_MSYS_DF_Click:
    Exit Sub

Err_Cmd_Cancella_FORM_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_Cancella_FORM_MSYS_DF_Click
End Sub





' ============================================================
' COMANDO: @CMD@MODULI@Importa
' ============================================================
Public Sub Cmd_IMPORTA_MODULI_MSYS_DF_Click()
    On Error GoTo Err_Cmd_IMPORTA_MODULI_MSYS_DF_Click
    
    Importa_MODULI_MSYS_DF_PFunct

Exit_Cmd_IMPORTA_MODULI_MSYS_DF_Click:
    Exit Sub

Err_Cmd_IMPORTA_MODULI_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_IMPORTA_MODULI_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: @CMD@MODULI@CANCELLA
' ============================================================
Public Sub Cmd_CANCELLA_MODULI_MSYS_DF_Click()
    On Error GoTo Err_Cmd_CANCELLA_MODULI_MSYS_DF_Click
    
    Call CancellaTuttiModuli_FORM_MSYS_DF

Exit_Cmd_CANCELLA_MODULI_MSYS_DF_Click:
    Exit Sub

Err_Cmd_CANCELLA_MODULI_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_CANCELLA_MODULI_MSYS_DF_Click
End Sub





' ============================================================
' COMANDO: @CMD@REPORT@Importa
' ============================================================
Public Sub Cmd_IMPORTA_REPORT_MSYS_DF_Click()
    On Error GoTo Err_Cmd_IMPORTA_REPORT_MSYS_DF_Click
    
    Importa_REPORT_MSYS_DF_PFunct

Exit_Cmd_IMPORTA_REPORT_MSYS_DF_Click:
    Exit Sub

Err_Cmd_IMPORTA_REPORT_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_IMPORTA_REPORT_MSYS_DF_Click
End Sub


' ============================================================
' COMANDO: @CMD@REPORT@CANCELLA
' ============================================================
Public Sub Cmd_CANCELLA_REPORT_MSYS_DF_Click()
    On Error GoTo Err_Cmd_CANCELLA_REPORT_MSYS_DF_Click
    
    Call CancellaTutti_REPORT_MSYS_DF

Exit_Cmd_CANCELLA_REPORT_MSYS_DF_Click:
    Exit Sub

Err_Cmd_CANCELLA_REPORT_MSYS_DF_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_CANCELLA_REPORT_MSYS_DF_Click
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
' @02 FUNZIONE Importa_LINK_TABELLE_MSYS_DF_PFunct
' SCOPO:
'   Collega tutte le tabelle definite in GetTabellaPaths.
'   - Controlla NomeTabella e PercorsoFile
'   - Scrive un log dettagliato
' ============================================================
Public Function Importa_LINK_TABELLE_MSYS_DF_PFunct()
    On Error GoTo Importa_LINK_TABELLE_MSYS_DF_PFunct_Err

    Dim TabellaPath As Collection
    Set TabellaPath = GetTabellaPaths_LINK_TABELLE_MSYS_DF()
    
    Dim NomeTabella As String
    Dim PercorsoFile As String
    Dim i As Long
    Dim LogFile As Integer
    Dim LogFileName As String
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long

    If TabellaPath.count = 0 Then
        MsgBox "La collection delle tabelle è vuota. Nessuna operazione effettuata.", vbInformation
        Exit Function
    End If

    ' Creazione file log
    LogFileName = CurrentProject.Path & "\CollegaTabelle_LINK_TABELLE_MSYS_DF_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile

    Print #LogFile, "Log esecuzione funzione: CollegaTabelle_LINK_TABELLE_MSYS_DF"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ConteggioOk = 0
    ConteggioErr = 0

    ' Ciclo su tutte le tabelle
    For i = 1 To TabellaPath.count
        Dim Item As Variant
        Item = TabellaPath(i)
        NomeTabella = Item(0)
        PercorsoFile = Item(1)

        ' Controllo valori vuoti
        If Len(NomeTabella) = 0 Or Len(PercorsoFile) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome tabella o percorso vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo ProssimaTabella
        End If

        ' Controllo percorso file
        If Not FileEsiste(PercorsoFile) Then
            Print #LogFile, Now & " - ERRORE: Percorso file NON valido per " & NomeTabella & " -> " & PercorsoFile
            ConteggioErr = ConteggioErr + 1
            GoTo ProssimaTabella
        End If

        ' Elimina tabella se già esiste, poi crea collegamento
        On Error Resume Next
        EliminaTabellaSeEsiste_LINK_TABELLE_MSYS_DF NomeTabella
        DoCmd.TransferDatabase acLink, "Microsoft Access", PercorsoFile, acTable, NomeTabella, NomeTabella, False

        If Err.number = 0 Then
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

    ' Riepilogo log
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "TOTALE TABELLE COLLEGATE: " & ConteggioOk
    Print #LogFile, "TOTALE ERRORI/SKIPPED: " & ConteggioErr
    Close #LogFile

    MsgBox "Operazione LINK TABELLE completata!" & vbCrLf & _
           "Tabelle collegate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log salvato in: " & LogFileName, vbInformation

Importa_LINK_TABELLE_MSYS_DF_PFunct_Exit:
    Exit Function

Importa_LINK_TABELLE_MSYS_DF_PFunct_Err:
    MsgBox "Errore: " & Err.Description, vbExclamation, "Errore di collegamento tabelle"
    Resume Importa_LINK_TABELLE_MSYS_DF_PFunct_Exit
End Function




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
    If TabellaPath.count = 0 Then
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
    Dim LogFile As Integer
    Dim LogFileName As String

    ' === Inizializza log ===
    LogFileName = CurrentProject.Path & "\CANC_TutteTabelleCollegate_PFunct_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile

    Print #LogFile, "Log esecuzione funzione: CancellaTutteTabelleCollegate_LINK_TABELLE_MSYS_DF_PFunct"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ConteggioOk = 0
    ConteggioErr = 0

    ' === Ciclo sugli elementi della collection ===
    For i = 1 To TabellaPath.count
        Item = TabellaPath(i) ' ogni elemento è un array con NomeTabella e PercorsoFile

        ' --- Estrazione sicura in variabili tipizzate (MODIFICA) ---
        NomeTabella = Trim$(CStr(Item(0) & ""))
        PercorsoFile = Trim$(CStr(Item(1) & ""))
        
        
        ' --- Controllo nome tabella vuoto (MODIFICA: rimosso percorso file) ---
        If Len(NomeTabella) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome tabella vuota"
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
            If Err.number = 0 Then
                Print #LogFile, Now & " - OK: Eliminata " & NomeTabella
                ConteggioOk = ConteggioOk + 1
            Else
                Print #LogFile, Now & " - ERRORE: " & NomeTabella & " -> " & Err.Description
                ConteggioErr = ConteggioErr + 1
                Err.Clear
            End If
            On Error GoTo 0
        Else
            Print #LogFile, Now & " - NON TROVATA: " & NomeTabella
            ConteggioErr = ConteggioErr + 1
        End If

ProssimaTabella:
    Next i

    ' === Riepilogo log ===
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "TOTALE TABELLE ELIMINATE: " & ConteggioOk
    Print #LogFile, "TOTALE ERRORI/SKIPPED: " & ConteggioErr
    Close #LogFile

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
'//                                         @IMPORTA@QUERY  O @CANCELLA@ QUERY                       *** INIZIO ***                 //
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
' FUNZIONE Importa_QUERY_MSYS_DF_PFunct
' SCOPO:
'   Collega tutte le query definite in GetQueryPaths_QUERY_MSYS_DF
'   - Scrive log dettagliato
' ============================================================
Public Function Importa_QUERY_MSYS_DF_PFunct()
    On Error GoTo ErrHandler

    Dim QueryPath As Collection
    Set QueryPath = GetQueryPaths_QUERY_MSYS_DF()
    
    Dim NomeQuery As String
    Dim PercorsoFile As String
    Dim i As Long
    Dim LogFile As Integer
    Dim LogFileName As String
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long

    If QueryPath.count = 0 Then
        MsgBox "Nessuna query da collegare.", vbInformation
        Exit Function
    End If
    
    ' File log
    LogFileName = CurrentProject.Path & "\CollegaQuery_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile
    Print #LogFile, "Log esecuzione funzione: Importa_QUERY_MSYS_DF_PFunct"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ' Ciclo
    For i = 1 To QueryPath.count
        Dim Item As Variant
        Item = QueryPath(i)
        NomeQuery = Item(0)
        PercorsoFile = Item(1)
        
        If Len(NomeQuery) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome query vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        If Not FileEsiste(PercorsoFile) Then
            Print #LogFile, Now & " - ERRORE: File non trovato -> " & PercorsoFile
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

        If Err.number = 0 Then
            Print #LogFile, Now & " - OK: " & NomeQuery
            ConteggioOk = ConteggioOk + 1
        Else
            Print #LogFile, Now & " - ERRORE: " & NomeQuery & " -> " & Err.Description
            ConteggioErr = ConteggioErr + 1
            Err.Clear
        End If
        On Error GoTo 0
Prossimo:
    Next i
    
    ' Riepilogo
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "Totale collegate: " & ConteggioOk
    Print #LogFile, "Totale errori/skipped: " & ConteggioErr
    Close #LogFile
    
     ' === aggiorna riquadro oggetti ===
    Call RefreshDatabaseView

    MsgBox "Operazione IMPORTA QUERY completata!" & vbCrLf & _
           "Collegate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log: " & LogFileName, vbInformation
    Exit Function

ErrHandler:
    MsgBox "Errore in Importa_QUERY_MSYS_DF_PFunct: " & Err.Description, vbExclamation
End Function


' ============================================================
' SUB CancellaQueryCollegate_QUERY_MSYS_DF
' SCOPO:
'   Cancella tutte le query collegate definite in GetQueryPaths_QUERY_MSYS_DF
' ============================================================
Public Sub CancellaQueryCollegate_QUERY_MSYS_DF()
    Dim QueryPath As Collection
    Set QueryPath = GetQueryPaths_QUERY_MSYS_DF()
    
    If QueryPath.count = 0 Then Exit Sub
    
    Dim i As Long
    Dim Item As Variant
    Dim NomeQuery As String
    
    For i = 1 To QueryPath.count
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
    
    If QueryPath.count = 0 Then
        MsgBox "Nessuna query da cancellare.", vbInformation
        Exit Function
    End If
    
    Dim i As Long, Item As Variant
    Dim NomeQuery As String
    Dim ConteggioOk As Long, ConteggioErr As Long
    Dim LogFile As Integer, LogFileName As String
    
    LogFileName = CurrentProject.Path & "\CancellaQuery_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile
    
    Print #LogFile, "Log esecuzione funzione: CancellaTutte_QUERY_MSYS_DF"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"
    
    For i = 1 To QueryPath.count
        Item = QueryPath(i)
        NomeQuery = Trim$(CStr(Item(0)))
        
        If Len(NomeQuery) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        If QueryEsiste_QUERY_MSYS_DF(NomeQuery) Then
            EliminaQuerySeEsiste_QUERY_MSYS_DF NomeQuery
            If Not QueryEsiste_QUERY_MSYS_DF(NomeQuery) Then
                Print #LogFile, Now & " - OK: Eliminata " & NomeQuery
                ConteggioOk = ConteggioOk + 1
            Else
                Print #LogFile, Now & " - ERRORE: Non eliminata " & NomeQuery
                ConteggioErr = ConteggioErr + 1
            End If
        Else
            Print #LogFile, Now & " - NON TROVATA: " & NomeQuery
            ConteggioErr = ConteggioErr + 1
        End If
Prossimo:
    Next i
    
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "Totale eliminate: " & ConteggioOk
    Print #LogFile, "Totale errori/skipped: " & ConteggioErr
    Close #LogFile
    
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
'//                                         @IMPORTA@QUERY  O @CANCELLA@ QUERY                       *** FINE ***                   //
'//*********************************************************************************************************************************//




'//*********************************************************************************************************************************//
'//                                         @IMPORTA@FORM  O @CANCELLA@FORM                         *** INIZIO ***                 //
'//*********************************************************************************************************************************//


' ============================================================
' FUNZIONE GetFormPaths_FORM_MSYS_DF
' SCOPO:
'   Restituisce una collection di Array(NomeForm, PercorsoFile)
'   che definisce le form da collegare dal DB sorgente.
' ============================================================
Public Function GetFormPaths_FORM_MSYS_DF() As Collection
    On Error GoTo GestErrore
    
    Dim FormPath As New Collection
    Dim PercorsoFile As String
    
    ' === Percorso del file sorgente contenente le form ===
    PercorsoFile = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb"
    
    If Not FileEsiste(PercorsoFile) Then
        MsgBox "File mancante: " & vbCrLf & PercorsoFile, vbExclamation
        Set GetFormPaths_FORM_MSYS_DF = FormPath
        Exit Function
    End If
    
    ' === Elenco form da collegare ===
    FormPath.Add Array("Msys_DF01_Frm01_}----------------------------------------------@", PercorsoFile)
    FormPath.Add Array("Msys_DF01_Frm01_M01_GE_COMANDI", PercorsoFile)
    FormPath.Add Array("Msys_DF01_Frm01_S01_COMANDI", PercorsoFile)
    
    Set GetFormPaths_FORM_MSYS_DF = FormPath
    
    Exit Function

GestErrore:
    MsgBox "Errore in GetFormPaths_FORM_MSYS_DF: " & Err.Description, vbExclamation, "Errore"
    Set GetFormPaths_FORM_MSYS_DF = Nothing
End Function


' ============================================================
' FUNZIONE Importa_FORM_MSYS_DF_PFunct
' SCOPO:
'   Collega tutte le form definite in GetFormPaths_FORM_MSYS_DF
'   - Scrive log dettagliato
' ============================================================
Public Function Importa_FORM_MSYS_DF_PFunct()
    On Error GoTo ErrHandler

    Dim FormPath As Collection
    Set FormPath = GetFormPaths_FORM_MSYS_DF()
    
    Dim NomeForm As String
    Dim PercorsoFile As String
    Dim i As Long
    Dim LogFile As Integer
    Dim LogFileName As String
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long

    If FormPath.count = 0 Then
        MsgBox "Nessuna form da collegare.", vbInformation
        Exit Function
    End If
    
    ' File log
    LogFileName = CurrentProject.Path & "\CollegaForm_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile
    Print #LogFile, "Log esecuzione funzione: Importa_FORM_MSYS_DF_PFunct"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ' Ciclo
    For i = 1 To FormPath.count
        Dim Item As Variant
        Item = FormPath(i)
        NomeForm = Item(0)
        PercorsoFile = Item(1)
        
        If Len(NomeForm) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome form vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        If Not FileEsiste(PercorsoFile) Then
            Print #LogFile, Now & " - ERRORE: File non trovato -> " & PercorsoFile
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        ' Elimina se già esiste
        EliminaFormSeEsiste_FORM_MSYS_DF NomeForm
        
        ' Collega
        On Error Resume Next
        
        '//@cmd@importa@FORM
        DoCmd.TransferDatabase _
            TransferType:=acImport, _
            DatabaseType:="Microsoft Access", _
            DatabaseName:=PercorsoFile, _
            ObjectType:=acForm, _
            Source:=NomeForm, _
            Destination:=NomeForm, _
            StructureOnly:=False

        If Err.number = 0 Then
            Print #LogFile, Now & " - OK: " & NomeForm
            ConteggioOk = ConteggioOk + 1
        Else
            Print #LogFile, Now & " - ERRORE: " & NomeForm & " -> " & Err.Description
            ConteggioErr = ConteggioErr + 1
            Err.Clear
        End If
        On Error GoTo 0
Prossimo:
    Next i
    
    ' Riepilogo
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "Totale collegate: " & ConteggioOk
    Print #LogFile, "Totale errori/skipped: " & ConteggioErr
    Close #LogFile
    
     ' === aggiorna riquadro oggetti ===
    Call RefreshDatabaseView

    MsgBox "Operazione IMPORTA FORM completata!" & vbCrLf & _
           "Collegate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log: " & LogFileName, vbInformation
    Exit Function

ErrHandler:
    MsgBox "Errore in Importa_FORM_MSYS_DF_PFunct: " & Err.Description, vbExclamation
End Function


' ============================================================
' SUB CancellaSingola_FORM_MSYS_DF
' SCOPO:
'   Cancella singola form
' ============================================================
Public Sub CancellaSingola_FORM_MSYS_DF()
    Dim FormPath As Collection
    Set FormPath = GetFormPaths_FORM_MSYS_DF()
    
    If FormPath.count = 0 Then Exit Sub
    
    Dim i As Long
    Dim Item As Variant
    Dim NomeForm As String
    
    For i = 1 To FormPath.count
        Item = FormPath(i)
        NomeForm = CStr(Item(0))
        
        If Len(NomeForm) > 0 Then
            EliminaFormSeEsiste_FORM_MSYS_DF NomeForm
        End If
    Next i
    
    MsgBox "Cancellazione FORM completata!", vbInformation
    
    '//TODO : manca la gestione errori routine
End Sub


' ============================================================
' FUNZIONE CancellaTutte_FORM_MSYS_DF
' SCOPO:
'   Cancella tutte le form collegate definite in GetFormPaths_FORM_MSYS_DF
'   - Scrive log dettagliato
' ============================================================
Public Function CancellaTutte_FORM_MSYS_DF()
    On Error GoTo ErrHandler
    
    Dim FormPath As Collection
    Set FormPath = GetFormPaths_FORM_MSYS_DF()
    
    If FormPath.count = 0 Then
        MsgBox "Nessuna form da cancellare.", vbInformation
        Exit Function
    End If
    
    Dim i As Long, Item As Variant
    Dim NomeForm As String
    Dim ConteggioOk As Long, ConteggioErr As Long
    Dim LogFile As Integer, LogFileName As String
    
    LogFileName = CurrentProject.Path & "\CancellaForm_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile
    
    Print #LogFile, "Log esecuzione funzione: CancellaTutte_FORM_MSYS_DF"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"
    
    For i = 1 To FormPath.count
        Item = FormPath(i)
        NomeForm = Trim$(CStr(Item(0)))
        
        If Len(NomeForm) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        If FormEsiste_FORM_MSYS_DF(NomeForm) Then
            EliminaFormSeEsiste_FORM_MSYS_DF NomeForm
            If Not FormEsiste_FORM_MSYS_DF(NomeForm) Then
                Print #LogFile, Now & " - OK: Eliminata " & NomeForm
                ConteggioOk = ConteggioOk + 1
            Else
                Print #LogFile, Now & " - ERRORE: Non eliminata " & NomeForm
                ConteggioErr = ConteggioErr + 1
            End If
        Else
            Print #LogFile, Now & " - NON TROVATA: " & NomeForm
            ConteggioErr = ConteggioErr + 1
        End If
Prossimo:
    Next i
    
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "Totale eliminate: " & ConteggioOk
    Print #LogFile, "Totale errori/skipped: " & ConteggioErr
    Close #LogFile
    
    Call RefreshDatabaseView
    MsgBox "Cancellazione FORM completata!" & vbCrLf & _
           "Eliminate: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log: " & LogFileName, vbInformation
    Exit Function
    
ErrHandler:
    MsgBox "Errore in CancellaTutte_FORM_MSYS_DF: " & Err.Description, vbExclamation
End Function


' ============================================================
' SUB EliminaFormSeEsiste_FORM_MSYS_DF
' SCOPO:
'   Elimina una form dal database corrente se esiste
' ============================================================
Public Sub EliminaFormSeEsiste_FORM_MSYS_DF(NomeForm As String)
    On Error Resume Next
    DoCmd.DeleteObject acForm, NomeForm
    On Error GoTo 0
End Sub


' ============================================================
' FUNZIONE FormEsiste_FORM_MSYS_DF
' SCOPO:
'   Verifica se una form esiste nel DB corrente
' ============================================================
Public Function FormEsiste_FORM_MSYS_DF(NomeForm As String) As Boolean
    Dim obj As AccessObject
    FormEsiste_FORM_MSYS_DF = False
    For Each obj In CurrentProject.AllForms
        If obj.Name = NomeForm Then
            FormEsiste_FORM_MSYS_DF = True
            Exit For
        End If
    Next obj
End Function


'//*********************************************************************************************************************************//
'//                                         @IMPORTA@FORM  O @CANCELLA@ FORM                         *** FINE ***                   //
'//*********************************************************************************************************************************//





'//*********************************************************************************************************************************//
'//                                         @IMPORTA@MODULI  O @CANCELLA@MODULI                         *** INIZIO ***              //
'//*********************************************************************************************************************************//


' ============================================================
' FUNZIONE GetModulePaths_MODULI_MSYS_DF
' SCOPO:
'   Restituisce una collection di Array(NomeModulo, PercorsoFile)
'   che definisce i moduli da collegare dal DB sorgente.
' ============================================================
Public Function GetModulePaths_MODULI_MSYS_DF() As Collection
    On Error GoTo GestErrore
    
    Dim ModulePath As New Collection
    Dim PercorsoFile As String
    
    ' === Percorso del file sorgente contenente i moduli ===
    PercorsoFile = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb"
    
    If Not FileEsiste(PercorsoFile) Then
        MsgBox "File mancante: " & vbCrLf & PercorsoFile, vbExclamation
        Set GetModulePaths_MODULI_MSYS_DF = ModulePath
        Exit Function
    End If
    
    ' === Elenco moduli da collegare ===
    ModulePath.Add Array("CLASSE_ATTI_Tb01_Gestione", PercorsoFile)
    ModulePath.Add Array("Msys_DF01_Mdl00_}----------------------------------------------@", PercorsoFile)
    ModulePath.Add Array("Msys_DF01_Mdl01_COLLEGA_COMANDI", PercorsoFile)
    
    Set GetModulePaths_MODULI_MSYS_DF = ModulePath
    
    Exit Function

GestErrore:
    MsgBox "Errore in GetModulePaths_MODULI_MSYS_DF: " & Err.Description, vbExclamation, "Errore"
    Set GetModulePaths_MODULI_MSYS_DF = Nothing
End Function


' ============================================================
' FUNZIONE Importa_MODULI_MSYS_DF_PFunct
' SCOPO:
'   Collega tutti i moduli definiti in GetModulePaths_MODULI_MSYS_DF
'   - Scrive log dettagliato
' ============================================================
Public Function Importa_MODULI_MSYS_DF_PFunct()
    On Error GoTo ErrHandler

    Dim ModulePath As Collection
    Set ModulePath = GetModulePaths_MODULI_MSYS_DF()
    
    Dim NomeModulo As String
    Dim PercorsoFile As String
    Dim i As Long
    Dim LogFile As Integer
    Dim LogFileName As String
    Dim ConteggioOk As Long
    Dim ConteggioErr As Long

    If ModulePath.count = 0 Then
        MsgBox "Nessun modulo da collegare.", vbInformation
        Exit Function
    End If
    
    ' File log
    LogFileName = CurrentProject.Path & "\CollegaModulo_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile
    Print #LogFile, "Log esecuzione funzione: Importa_MODULI_MSYS_DF_PFunct"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"

    ' Ciclo
    For i = 1 To ModulePath.count
        Dim Item As Variant
        Item = ModulePath(i)
        NomeModulo = Item(0)
        PercorsoFile = Item(1)
        
        If Len(NomeModulo) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome modulo vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        If Not FileEsiste(PercorsoFile) Then
            Print #LogFile, Now & " - ERRORE: File non trovato -> " & PercorsoFile
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        ' Elimina se già esiste
        EliminaModuloSeEsiste_FORM_MSYS_DF NomeModulo
        
        ' Collega
        On Error Resume Next
        
        '//@cmd@importa@MODULI
        DoCmd.TransferDatabase _
            TransferType:=acImport, _
            DatabaseType:="Microsoft Access", _
            DatabaseName:=PercorsoFile, _
            ObjectType:=acModule, _
            Source:=NomeModulo, _
            Destination:=NomeModulo, _
            StructureOnly:=False

        If Err.number = 0 Then
            Print #LogFile, Now & " - OK: " & NomeModulo
            ConteggioOk = ConteggioOk + 1
        Else
            Print #LogFile, Now & " - ERRORE: " & NomeModulo & " -> " & Err.Description
            ConteggioErr = ConteggioErr + 1
            Err.Clear
        End If
        On Error GoTo 0
Prossimo:
    Next i
    
    ' Riepilogo
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "Totale collegati: " & ConteggioOk
    Print #LogFile, "Totale errori/skipped: " & ConteggioErr
    Close #LogFile
    
     ' === aggiorna riquadro oggetti ===
    Call RefreshDatabaseView

    MsgBox "Operazione IMPORTA MODULI completata!" & vbCrLf & _
           "Collegati: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log: " & LogFileName, vbInformation
    Exit Function

ErrHandler:
    MsgBox "Errore in Importa_MODULI_MSYS_DF_PFunct: " & Err.Description, vbExclamation
End Function


' ============================================================
' SUB CancellaModuliCollegati_FORM_MSYS_DF
' SCOPO:
'   Cancella tutti i moduli collegati definiti in GetModulePaths_MODULI_MSYS_DF
' ============================================================
Public Sub CancellaModuliCollegati_FORM_MSYS_DF()
    Dim ModulePath As Collection
    Set ModulePath = GetModulePaths_MODULI_MSYS_DF()
    
    If ModulePath.count = 0 Then Exit Sub
    
    Dim i As Long
    Dim Item As Variant
    Dim NomeModulo As String
    
    For i = 1 To ModulePath.count
        Item = ModulePath(i)
        NomeModulo = CStr(Item(0))
        
        If Len(NomeModulo) > 0 Then
            EliminaModuloSeEsiste_FORM_MSYS_DF NomeModulo
        End If
    Next i
    
    MsgBox "Cancellazione MODULI completata!", vbInformation
    
    '//TODO : manca la gestione errori routine
End Sub


' ============================================================
' FUNZIONE CancellaTuttiModuli_FORM_MSYS_DF
' SCOPO:
'   Cancella tutti i moduli collegati definiti in GetModulePaths_MODULI_MSYS_DF
'   - Scrive log dettagliato
' ============================================================
Public Function CancellaTuttiModuli_FORM_MSYS_DF()
    On Error GoTo ErrHandler
    
    Dim ModulePath As Collection
    Set ModulePath = GetModulePaths_MODULI_MSYS_DF()
    
    If ModulePath.count = 0 Then
        MsgBox "Nessun modulo da cancellare.", vbInformation
        Exit Function
    End If
    
    Dim i As Long, Item As Variant
    Dim NomeModulo As String
    Dim ConteggioOk As Long, ConteggioErr As Long
    Dim LogFile As Integer, LogFileName As String
    
    LogFileName = CurrentProject.Path & "\CancellaModulo_LOG.txt"
    LogFile = FreeFile
    Open LogFileName For Output As #LogFile
    
    Print #LogFile, "Log esecuzione funzione: CancellaTuttiModuli_FORM_MSYS_DF"
    Print #LogFile, "Data/Ora: " & Now()
    Print #LogFile, "--------------------------------------------"
    
    For i = 1 To ModulePath.count
        Item = ModulePath(i)
        NomeModulo = Trim$(CStr(Item(0)))
        
        If Len(NomeModulo) = 0 Then
            Print #LogFile, Now & " - SKIPPED: Nome vuoto"
            ConteggioErr = ConteggioErr + 1
            GoTo Prossimo
        End If
        
        If ControlloEsistenza_MODULI_MSYS_DF(NomeModulo) Then
            EliminaModuloSeEsiste_FORM_MSYS_DF NomeModulo
            If Not ControlloEsistenza_MODULI_MSYS_DF(NomeModulo) Then
                Print #LogFile, Now & " - OK: Eliminato " & NomeModulo
                ConteggioOk = ConteggioOk + 1
            Else
                Print #LogFile, Now & " - ERRORE: Non eliminato " & NomeModulo
                ConteggioErr = ConteggioErr + 1
            End If
        Else
            Print #LogFile, Now & " - NON TROVATO: " & NomeModulo
            ConteggioErr = ConteggioErr + 1
        End If
Prossimo:
    Next i
    
    Print #LogFile, "--------------------------------------------"
    Print #LogFile, "Totale eliminati: " & ConteggioOk
    Print #LogFile, "Totale errori/skipped: " & ConteggioErr
    Close #LogFile
    
    Call RefreshDatabaseView
    MsgBox "Cancellazione MODULI completata!" & vbCrLf & _
           "Eliminati: " & ConteggioOk & vbCrLf & _
           "Errori/Skipped: " & ConteggioErr & vbCrLf & _
           "Log: " & LogFileName, vbInformation
    Exit Function
    
ErrHandler:
    MsgBox "Errore in CancellaTuttiModuli_FORM_MSYS_DF: " & Err.Description, vbExclamation
End Function


' ============================================================
' SUB EliminaModuloSeEsiste_FORM_MSYS_DF
' SCOPO:
'   Elimina un modulo dal database corrente se esiste
' ============================================================
Public Sub EliminaModuloSeEsiste_FORM_MSYS_DF(NomeModulo As String)
    On Error Resume Next
    DoCmd.DeleteObject acModule, NomeModulo
    On Error GoTo 0
End Sub


' ============================================================
' FUNZIONE ControlloEsistenza_MODULI_MSYS_DF
' SCOPO:
'   Verifica se un modulo esiste nel DB corrente
' ============================================================
Public Function ControlloEsistenza_MODULI_MSYS_DF(NomeModulo As String) As Boolean
    Dim obj As AccessObject
    ControlloEsistenza_MODULI_MSYS_DF = False
    For Each obj In CurrentProject.AllModules
        If obj.Name = NomeModulo Then
            ControlloEsistenza_MODULI_MSYS_DF = True
            Exit For
        End If
    Next obj
End Function


'//*********************************************************************************************************************************//
'//                                         @IMPORTA@MODULI  O @CANCELLA@ MODULO                         *** FINE ***               //
'//*********************************************************************************************************************************//


'//*********************************************************************************************************************************//
'//                                         @IMPORTA@REPORT  O @CANCELLA@REPORT                       *** INIZIO ***                //
'//*********************************************************************************************************************************//

    ' ============================================================
    ' FUNZIONE GetReportPaths_REPORT_MSYS_DF
    ' SCOPO:
    '   Restituisce una collection di Array(NomeReport, PercorsoFile)
    '   che definisce i report da collegare dal DB sorgente.
    ' ============================================================
    Public Function GetReportPaths_REPORT_MSYS_DF() As Collection
        On Error GoTo GestErrore
        
        Dim ReportPath As New Collection
        Dim PercorsoFile As String
        
        ' === Percorso del file sorgente contenente i report ===
        PercorsoFile = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb"
        
        If Not FileEsiste(PercorsoFile) Then
            MsgBox "File mancante: " & vbCrLf & PercorsoFile, vbExclamation
            Set GetReportPaths_REPORT_MSYS_DF = ReportPath
            Exit Function
        End If
        
        ' === Elenco report da collegare ===
        ReportPath.Add Array("Msys_DF01_Rpt00_}----------------------------------------------@", PercorsoFile)
        ReportPath.Add Array("Msys_DF01_Rpt01_DA_DEFINIRE", PercorsoFile)
        'ReportPath.Add Array("Msys_DF01_Rpt01_S01_COMANDI", PercorsoFile)
        
        Set GetReportPaths_REPORT_MSYS_DF = ReportPath
        Exit Function

GestErrore:
        MsgBox "Errore in GetReportPaths_REPORT_MSYS_DF: " & Err.Description, vbExclamation, "Errore"
        Set GetReportPaths_REPORT_MSYS_DF = Nothing
    End Function


    ' ============================================================
    ' FUNZIONE Importa_REPORT_MSYS_DF_PFunct
    ' SCOPO:
    '   Collega tutti i report definiti in GetReportPaths_REPORT_MSYS_DF
    '   - Scrive log dettagliato
    ' ============================================================
    Public Function Importa_REPORT_MSYS_DF_PFunct()
        On Error GoTo ErrHandler

        Dim ReportPath As Collection
        Set ReportPath = GetReportPaths_REPORT_MSYS_DF()
        
        Dim NomeReport As String
        Dim PercorsoFile As String
        Dim i As Long
        Dim LogFile As Integer
        Dim LogFileName As String
        Dim ConteggioOk As Long
        Dim ConteggioErr As Long

        If ReportPath.count = 0 Then
            MsgBox "Nessun report da collegare.", vbInformation
            Exit Function
        End If
        
        ' File log
        LogFileName = CurrentProject.Path & "\CollegaReport_LOG.txt"
        LogFile = FreeFile
        Open LogFileName For Output As #LogFile
        Print #LogFile, "Log esecuzione funzione: Importa_REPORT_MSYS_DF_PFunct"
        Print #LogFile, "Data/Ora: " & Now()
        Print #LogFile, "--------------------------------------------"

        ' Ciclo
        For i = 1 To ReportPath.count
            Dim Item As Variant
            Item = ReportPath(i)
            NomeReport = Item(0)
            PercorsoFile = Item(1)
            
            If Len(NomeReport) = 0 Then
                Print #LogFile, Now & " - SKIPPED: Nome report vuoto"
                ConteggioErr = ConteggioErr + 1
                GoTo Prossimo
            End If
            
            If Not FileEsiste(PercorsoFile) Then
                Print #LogFile, Now & " - ERRORE: File non trovato -> " & PercorsoFile
                ConteggioErr = ConteggioErr + 1
                GoTo Prossimo
            End If
            
            ' Elimina se già esiste
            EliminaReportSeEsiste_REPORT_MSYS_DF NomeReport
            
            ' Collega
            On Error Resume Next
            DoCmd.TransferDatabase _
                TransferType:=acImport, _
                DatabaseType:="Microsoft Access", _
                DatabaseName:=PercorsoFile, _
                ObjectType:=acReport, _
                Source:=NomeReport, _
                Destination:=NomeReport, _
                StructureOnly:=False
            On Error GoTo 0

            ' Controllo effettivo esistenza
            If ReportEsiste_REPORT_MSYS_DF(NomeReport) Then
                Print #LogFile, Now & " - OK: " & NomeReport
                ConteggioOk = ConteggioOk + 1
            Else
                Print #LogFile, Now & " - ERRORE: " & NomeReport
                ConteggioErr = ConteggioErr + 1
            End If

Prossimo:
        Next i
        
        ' Riepilogo
        Print #LogFile, "--------------------------------------------"
        Print #LogFile, "Totale collegati: " & ConteggioOk
        Print #LogFile, "Totale errori/skipped: " & ConteggioErr
        Close #LogFile
        
         ' === aggiorna riquadro oggetti ===
        Call RefreshDatabaseView

        MsgBox "Operazione IMPORTA REPORT completata!" & vbCrLf & _
               "Collegati: " & ConteggioOk & vbCrLf & _
               "Errori/Skipped: " & ConteggioErr & vbCrLf & _
               "Log: " & LogFileName, vbInformation
        Exit Function

ErrHandler:
        MsgBox "Errore in Importa_REPORT_MSYS_DF_PFunct: " & Err.Description, vbExclamation
    End Function


    ' ============================================================
    ' SUB CancellaSingolo_REPORT_MSYS_DF
    ' SCOPO:
    '   Cancella tutti i report collegati definiti in GetReportPaths_REPORT_MSYS_DF
    ' ============================================================
    Public Sub CancellaSingolo_REPORT_MSYS_DF()
        Dim ReportPath As Collection
        Set ReportPath = GetReportPaths_REPORT_MSYS_DF()
        
        If ReportPath.count = 0 Then Exit Sub
        
        Dim i As Long
        Dim Item As Variant
        Dim NomeReport As String
        
        For i = 1 To ReportPath.count
            Item = ReportPath(i)
            NomeReport = CStr(Item(0))
            
            If Len(NomeReport) > 0 Then
                EliminaReportSeEsiste_REPORT_MSYS_DF NomeReport
            End If
        Next i
        
        MsgBox "Cancellazione REPORT completata!", vbInformation
        
        '//TODO : manca la gestione errori routine
    End Sub


    ' ============================================================
    ' FUNZIONE CancellaTutti_REPORT_MSYS_DF
    ' SCOPO:
    '   Cancella tutti i report collegati definiti in GetReportPaths_REPORT_MSYS_DF
    '   - Scrive log dettagliato
    ' ============================================================
    Public Function CancellaTutti_REPORT_MSYS_DF()
        On Error GoTo ErrHandler
        
        Dim ReportPath As Collection
        Set ReportPath = GetReportPaths_REPORT_MSYS_DF()
        
        If ReportPath.count = 0 Then
            MsgBox "Nessun report da cancellare.", vbInformation
            Exit Function
        End If
        
        Dim i As Long, Item As Variant
        Dim NomeReport As String
        Dim ConteggioOk As Long, ConteggioErr As Long
        Dim LogFile As Integer, LogFileName As String
        
        LogFileName = CurrentProject.Path & "\CancellaReport_LOG.txt"
        LogFile = FreeFile
        Open LogFileName For Output As #LogFile
        
        Print #LogFile, "Log esecuzione funzione: CancellaTutti_REPORT_MSYS_DF"
        Print #LogFile, "Data/Ora: " & Now()
        Print #LogFile, "--------------------------------------------"
        
        For i = 1 To ReportPath.count
            Item = ReportPath(i)
            NomeReport = Trim$(CStr(Item(0)))
            
            If Len(NomeReport) = 0 Then
                Print #LogFile, Now & " - SKIPPED: Nome vuoto"
                ConteggioErr = ConteggioErr + 1
                GoTo Prossimo
            End If
            
            If ReportEsiste_REPORT_MSYS_DF(NomeReport) Then
                EliminaReportSeEsiste_REPORT_MSYS_DF NomeReport
                If Not ReportEsiste_REPORT_MSYS_DF(NomeReport) Then
                    Print #LogFile, Now & " - OK: Eliminato " & NomeReport
                    ConteggioOk = ConteggioOk + 1
                Else
                    Print #LogFile, Now & " - ERRORE: Non eliminato " & NomeReport
                    ConteggioErr = ConteggioErr + 1
                End If
            Else
                Print #LogFile, Now & " - NON TROVATO: " & NomeReport
                ConteggioErr = ConteggioErr + 1
            End If
Prossimo:
        Next i
        
        Print #LogFile, "--------------------------------------------"
        Print #LogFile, "Totale eliminati: " & ConteggioOk
        Print #LogFile, "Totale errori/skipped: " & ConteggioErr
        Close #LogFile
        
        Call RefreshDatabaseView
        MsgBox "Cancellazione REPORT completata!" & vbCrLf & _
               "Eliminati: " & ConteggioOk & vbCrLf & _
               "Errori/Skipped: " & ConteggioErr & vbCrLf & _
               "Log: " & LogFileName, vbInformation
        Exit Function
        
ErrHandler:
        MsgBox "Errore in CancellaTutti_REPORT_MSYS_DF: " & Err.Description, vbExclamation
    End Function


    ' ============================================================
    ' SUB EliminaReportSeEsiste_REPORT_MSYS_DF
    ' SCOPO:
    '   Elimina un report dal database corrente se esiste
    ' ============================================================
    Public Sub EliminaReportSeEsiste_REPORT_MSYS_DF(NomeReport As String)
        On Error Resume Next
        DoCmd.DeleteObject acReport, NomeReport
        On Error GoTo 0
    End Sub


    ' ============================================================
    ' FUNZIONE ReportEsiste_REPORT_MSYS_DF
    ' SCOPO:
    '   Verifica se un report esiste nel DB corrente
    ' ============================================================
    Public Function ReportEsiste_REPORT_MSYS_DF(NomeReport As String) As Boolean
        Dim obj As AccessObject
        ReportEsiste_REPORT_MSYS_DF = False
        For Each obj In CurrentProject.AllReports
            If obj.Name = NomeReport Then
                ReportEsiste_REPORT_MSYS_DF = True
                Exit For
            End If
        Next obj
    End Function


'//*********************************************************************************************************************************//
'//                                         @IMPORTA@REPORT  O @CANCELLA@REPORT                       *** FINE ***                  //
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
    
    '//IMPOSTO LA VARIABILE
    Str1 = ""
    Str1 = "Gestione Oggetti MSys_DF01.mdb"
    ' --------------------------------------------------------------
    ' Inizializzazione delle variabili e impostazioni della form
    ' --------------------------------------------------------------
    Me.Caption = Str1   ' Imposta il titolo della finestra
    
    ' TITLO LABEL
    Me.LBL_TIT01_TXT.Caption = Str1


' --------------------------------------------------------------
' uscita + Gestione errori
' --------------------------------------------------------------


Exit Sub   ' Uscita normale dal metodo

ErrHandler:
    MsgBox "Errore durante il caricamento della form: " & Err.Description, _
           vbCritical, "Errore " & Err.number
    Resume Next   ' Continua l'esecuzione dopo l'errore (valuta Resume/Exit in base al caso)
End Sub








'//*********************************************************************************************************************************//
'//                                          @EVENTI@FORM A TUTTO IL PROGETTO                     *** FINE ***
'//*********************************************************************************************************************************//


