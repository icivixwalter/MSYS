Attribute VB_Name = "MSys_TB30_UTIL_Export_Mdl01_TABELLE_Link_RicreaCollegamento"
Option Compare Database
Option Explicit

'======================================================================
' MODULO    : LLPP_DF99_LINK_Mdl01_TABELLE_Link_RicreaCollegamento
' PROCEDURA : RICOSTRUICI_LINK_TABELLE_pf
' CODICE     : @SHADOW_SYSTEM_CATALOG
' AUTORE     : System
' DATA       : 2026
' VERSIONE   : 3.0 TB99 SELF-HEALING LINK ENGINE WITH LOGGING
'
'======================================================================
' SCOPO DEL MODULO
'======================================================================
' Questo modulo NON si limita più a ricreare collegamenti.
'
' È un motore di controllo e auto-riparazione dei link di Access
' basato sulla query centrale:
'
'       Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'
' La query TB99 rappresenta un repository logico dei collegamenti
' dell'intero sistema e viene utilizzata come fonte unica di verità.
'
' Il modulo genera un file di log dettagliato di ogni operazione
' per consentire audit, debug e tracciabilità delle riparazioni.
'
'======================================================================
' LOGICA OPERATIVA
'======================================================================
'
' Per ogni record della query TB99:
'
' 1) Legge il percorso completo del database (campo Database)
' 2) Legge il nome della tabella (campo Name)
' 3) Legge il tipo oggetto (campo Type, deve essere 6 = link)
' 4) Verifica che il path sia strutturalmente valido
' 5) Verifica esistenza file fisico del database di destinazione
' 6) APRE il database di destinazione
' 7) Verifica se la tabella esiste NEL DATABASE DI DESTINAZIONE
' 8) Se esiste, elimina il link corrente nel DB corrente
' 9) Ricostruisce il link dal DB di destinazione al DB corrente
' 10) Aggiorna lo stato nel campo Verifica_s con uno dei valori:
'
'       ? OK              -> link ricostruito e operativo
'       ? MISSING         -> file database non trovato
'       ? INVALID_PATH    -> path strutturalmente errato
'       ? NOT_REGISTERED  -> tabella non trovata nel DB sorgente O Type != 6
'       ? NOT_CHECKED     -> non ancora verificato
'
' 11) Traccia ogni operazione su file di log e finestra Immediate
' 12) Rilascia le risorse del DB di destinazione prima di passare al record successivo
'
'======================================================================
' DESCRIZIONE
'======================================================================
' Il modulo sincronizza i collegamenti tra:
'
'   - Database di destinazione (sorgente dati reale)
'   - Database corrente (dove vengono ricreati i link)
'   - Msys_TB99_LINK_Qry99_01_SELECT_Archivio (repository logico centrale)
'
' e mantiene coerenza tra filesystem, database sorgente e database corrente.
'
'======================================================================
' ARCHITETTURA CONCETTUALE
'======================================================================
'
'   Query TB99 = SINGLE SOURCE OF TRUTH (metadati link)
'   Database di destinazione = sorgente fisica delle tabelle
'   Database corrente = destinatario dei link
'   File di log = tracciamento operazioni
'   Modulo VBA = motore di verifica e ricostruzione link
'
'                     [Query TB99]
'                          |
'                          | (legge Name + Database)
'                          v
'                  [Database Destinazione]
'                          |
'                          | (apri -> verifica tabella -> chiudi)
'                          v
'                   [Ricostruzione Link]
'                          |
'                          | (elimina link esistente -> crea nuovo link)
'                          v
'                  [Database Corrente]
'                          |
'                          v
'                    [File di Log]
'
'======================================================================
' ISTRUZIONI PER L'UTILIZZO
'======================================================================
'
' 1) QUERY DI RIFERIMENTO
'    Deve esistere:
'       Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'
'    Campi minimi richiesti:
'       - Name          (nome tabella da collegare)
'       - Database      (percorso completo DB di destinazione)
'       - Type          (deve essere 6 per i link table)
'       - Verifica_s    (campo aggiornabile per lo stato)
'
' 2) CONFIGURAZIONE LOG
'    Variabili modificabili in testata:
'       MODULE_PATH     = Percorso base del modulo
'       LOG_FOLDER      = Cartella per i file di log (creata automaticamente)
'       LOG_FILE_NAME   = Nome base del file di log
'
' 3) AVVIO PROCEDURA
'    - F5 nel modulo
'    - oppure Call RICOSTRUICI_LINK_TABELLE_pf
'
' 4) OUTPUT DEL SISTEMA
'    - Aggiornamento campo Verifica_s nella query TB99
'    - Auto-riparazione link mancanti/corrotti
'    - File di log dettagliato in LOG_FOLDER
'    - Logging in finestra Immediate
'    - Statistiche finali di coerenza sistema
'
'======================================================================
' NOTE ARCHITETTURALI IMPORTANTI
'======================================================================
'
' ? Verifica lato sorgente: apre il DB di destinazione e controlla
'   l'esistenza reale della tabella prima di ricreare il link
' ? Chiude immediatamente il DB di destinazione dopo la verifica
'   per rilasciare le risorse e passare al record successivo
' ? Elimina preventivamente i link corrotti prima della ricostruzione
' ? Usa la query TB99 come livello di astrazione e controllo
' ? Genera file di log con timestamp per ogni esecuzione
' ? Traccia ogni singola operazione con dettaglio dello stato
' ? Permette migrazione sicura dei percorsi database
' ? Evita corruzione strutturale dei link Access
'
'======================================================================
' MARKER LOGICI DI FLUSSO (ALLINEATO AL CODICE REALE)
'======================================================================
' 01_INIZIO                         = Inizializzazione variabili / DB corrente
' 01A_CREAZIONE_LOG                 = Creazione file di log con timestamp
' 02_VERIFICA_ESISTENZA_QUERY       = Controllo esistenza query TB99
' 03_CONFERMA_UTENTE                = Conferma operazione da parte dell'utente
' 04_APERTURA_RECORDSET_QUERY       = Apertura query Msys_TB99_LINK_Qry99_01_SELECT_Archivio
' 05_INIZIALIZZAZIONE_STATISTICHE   = Reset contatori statistiche
' 06_CICLO_PRINCIPALE               = Loop principale su tutti i record
' 07_LETTURA_RECORD                 = Lettura Name, Database, Type
' 08_VALIDAZIONE_PATH               = Verifica Type=6 e path valido
' 09_APERTURA_DB_DESTINAZIONE       = OpenDatabase() sul campo Database
' 10_VERIFICA_TABELLA_SORGENTE      = Controllo esistenza tabella nel DB destinazione
' 10A_CHIUSURA_DB_DESTINAZIONE      = Close database e rilascio risorse
' 11_ELIMINAZIONE_LINK_PRECEDENTE   = DeleteLink() se esiste
' 12_RICOSTRUZIONE_LINK             = CreateLink() dal DB destinazione a DB corrente
' 13_AGGIORNAMENTO_RECORDSET        = Update campo Verifica_s
' 14_AGGIORNAMENTO_STATISTICHE      = Incremento contatori per stato
' 15_RECORD_SUCCESSIVO              = MoveNext
' 16_STATISTICHE_FINALI             = Report stato sistema su Debug e Log
' 17_CHIUSURA_RISORSE               = Cleanup DAO, FSO e file di log
' 18_MESSAGGIO_FINALE               = MsgBox con esito finale operazione
' 99_GESTIONE_ERRORI                = Error handler globale con logging
'======================================================================

'======================================================================
' VARIABILI PER LOGGING
'======================================================================
Private Const MODULE_PATH As String = "C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB99_LINK\"
Private LOG_FOLDER As String
Private Const LOG_FILE_NAME As String = "TB99_Link_Rebuild_Log.txt"

'======================================================================
' PROCEDURA PRINCIPALE
'======================================================================

Public Sub RICOSTRUICI_LINK_TABELLE_pf()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fso As Object
    Dim logFile As Object
    Dim logFilePath As String

    Dim sSQL As String
    Dim sName As String
    Dim sDatabasePath As String
    Dim sStatus As String
    Dim lType As Long

    Dim tableExistsInSource As Boolean
    Dim startTime As Date
    Dim endTime As Date

'======================================================================
' 01_INIZIO = Inizializzazione variabili / DB corrente
'======================================================================
    Set db = CurrentDb
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    startTime = Now

'======================================================================
' 01A_CREAZIONE_LOG = Creazione file di log con timestamp
'======================================================================
    ' Imposta il percorso di default
    LOG_FOLDER = MODULE_PATH & "LOGS\"
    
    ' Crea la cartella LOG se non esiste
    If Not fso.FolderExists(LOG_FOLDER) Then
        On Error Resume Next
        fso.CreateFolder LOG_FOLDER
        If Err.Number <> 0 Then
            ' Fallback su C:\TEMP\
            LOG_FOLDER = "C:\TEMP\"
            fso.CreateFolder LOG_FOLDER
        End If
        On Error GoTo 0
    End If
    
    ' Costruisce il percorso completo del file di log
    logFilePath = LOG_FOLDER & Format(startTime, "YYYYMMDD_HHMMSS") & "_" & LOG_FILE_NAME
    
    ' Crea il file di log
    Set logFile = fso.CreateTextFile(logFilePath, True)
    
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "INIZIO RICOSTRUZIONE LINK TB99")
    Call WriteLog(logFile, "Data/Ora avvio: " & Format(startTime, "dd/mm/yyyy HH:MM:SS"))
    Call WriteLog(logFile, "Database corrente: " & CurrentDb.Name)
    Call WriteLog(logFile, "File di log: " & logFilePath)
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "")

'======================================================================
' 02_VERIFICA_ESISTENZA_QUERY = Controllo esistenza query TB99
'======================================================================
    Call WriteLog(logFile, "[VERIFICA_QUERY] Controllo esistenza query Msys_TB99_LINK_Qry99_01_SELECT_Archivio...")
    
    If Not QueryExists("Msys_TB99_LINK_Qry99_01_SELECT_Archivio") Then
        Call WriteLog(logFile, "[ERRORE] Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio NON TROVATA!")
        Call WriteLog(logFile, "[FINE] Operazione terminata con errore")
        logFile.Close
        Set logFile = Nothing
        MsgBox "Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio non trovata", vbCritical
        Exit Sub
    End If
    
    Call WriteLog(logFile, "[VERIFICA_QUERY] Query trovata con successo.")
    Call WriteLog(logFile, "")

'======================================================================
' 03_CONFERMA_UTENTE = Conferma operazione da parte dell'utente
'======================================================================
    Call WriteLog(logFile, "[UTENTE] Richiesta conferma operazione...")
    
    If MsgBox("Avvio ricostruzione link tabelle da TB99?", vbYesNo) = vbNo Then
        Call WriteLog(logFile, "[UTENTE] Operazione ANNULLATA dall'utente.")
        Call WriteLog(logFile, "[FINE] Log chiuso.")
        logFile.Close
        Set logFile = Nothing
        Exit Sub
    End If
    
    Call WriteLog(logFile, "[UTENTE] Operazione CONFERMATA.")
    Call WriteLog(logFile, "")

'======================================================================
' 04_APERTURA_RECORDSET_QUERY = Apertura query Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'======================================================================
    Call WriteLog(logFile, "[APERTURA_QUERY] Apertura recordset su Msys_TB99_LINK_Qry99_01_SELECT_Archivio...")
    
    sSQL = "SELECT * FROM Msys_TB99_LINK_Qry99_01_SELECT_Archivio"
    Set rs = db.OpenRecordset(sSQL, dbOpenDynaset)
    
    Call WriteLog(logFile, "[APERTURA_QUERY] Recordset aperto. Record trovati: " & rs.RecordCount)
    Call WriteLog(logFile, "")

'======================================================================
' 05_INIZIALIZZAZIONE_STATISTICHE = Reset contatori statistiche
'======================================================================
    Dim nTot As Long: nTot = 0
    Dim nOk As Long: nOk = 0
    Dim nMissing As Long: nMissing = 0
    Dim nInvalidPath As Long: nInvalidPath = 0
    Dim nNotRegistered As Long: nNotRegistered = 0
    
    Call WriteLog(logFile, "[STATISTICHE] Contatori inizializzati.")
    Call WriteLog(logFile, "")

'======================================================================
' 06_CICLO_PRINCIPALE = Loop principale su tutti i record
'======================================================================
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "INIZIO CICLO ELABORAZIONE RECORD")
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "")
    
    Do While Not rs.EOF
        nTot = nTot + 1
        
        Call WriteLog(logFile, "----------------------------------------")
        Call WriteLog(logFile, "RECORD N. " & nTot)
        Call WriteLog(logFile, "----------------------------------------")

'======================================================================
' 07_LETTURA_RECORD = Lettura Name, Database, Type
'======================================================================
        sName = Nz(rs!Name, "")
        sDatabasePath = Nz(rs!Database, "")
        lType = Nz(rs!Type, 0)

        Call WriteLog(logFile, "[LETTURA_RECORD] Nome tabella: '" & sName & "'")
        Call WriteLog(logFile, "[LETTURA_RECORD] Database path: '" & sDatabasePath & "'")
        Call WriteLog(logFile, "[LETTURA_RECORD] Type: " & lType)
        
        sStatus = "NOT_CHECKED"

'======================================================================
' 08_VALIDAZIONE_PATH = Verifica Type=6 e path valido
'======================================================================
        Call WriteLog(logFile, "[VALIDAZIONE] Avvio validazione...")
        
        ' Verifica che sia un link table (Type = 6)
        If lType <> 6 Then
            sStatus = "NOT_REGISTERED"
            Call WriteLog(logFile, "[VALIDAZIONE] Type = " & lType & " (diverso da 6) -> NON è un link table")
            Call WriteLog(logFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
        ' Verifica path strutturalmente valido
        ElseIf Len(Trim(sDatabasePath)) = 0 Or InStr(1, sDatabasePath, ":\") = 0 Then
            sStatus = "INVALID_PATH"
            Call WriteLog(logFile, "[VALIDAZIONE] Path strutturalmente INVALIDO")
            Call WriteLog(logFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
        ' Verifica esistenza file fisico del database di destinazione
        ElseIf Not fso.FileExists(sDatabasePath) Then
            sStatus = "MISSING"
            Call WriteLog(logFile, "[VALIDAZIONE] File database NON TROVATO sul filesystem")
            Call WriteLog(logFile, "[VALIDAZIONE] Path cercato: '" & sDatabasePath & "'")
            Call WriteLog(logFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
        Else
            Call WriteLog(logFile, "[VALIDAZIONE] Path valido. File database ESISTE.")
            Call WriteLog(logFile, "[VALIDAZIONE] Inizio verifica lato SORGENTE...")
            
            ' ============================================================
            ' 09_APERTURA_DB_DESTINAZIONE = OpenDatabase() sul campo Database
            ' 10_VERIFICA_TABELLA_SORGENTE = Controllo esistenza tabella nel DB destinazione
            ' 10A_CHIUSURA_DB_DESTINAZIONE = Close database e rilascio risorse
            ' ============================================================
            
            ' Apre il database di destinazione e verifica esistenza tabella
            ' La funzione TableExistsInDatabase apre, verifica e CHIUDE il database
            tableExistsInSource = TableExistsInDatabase(sDatabasePath, sName, logFile)
            
            If tableExistsInSource Then
                Call WriteLog(logFile, "[VERIFICA_TABELLA] Tabella '" & sName & "' TROVATA nel database di destinazione.")
                
                ' ============================================================
                ' 11_ELIMINAZIONE_LINK_PRECEDENTE = DeleteLink() se esiste
                ' ============================================================
                
                If LinkExistsInMSysObjects(sName) Then
                    Call WriteLog(logFile, "[ELIMINAZIONE_LINK] Link esistente TROVATO in MSysObjects.")
                    Call WriteLog(logFile, "[ELIMINAZIONE_LINK] Eliminazione in corso...")
                    DeleteLink sName, logFile
                Else
                    Call WriteLog(logFile, "[ELIMINAZIONE_LINK] Link esistente NON trovato in MSysObjects.")
                End If
                
                ' ============================================================
                ' 12_RICOSTRUZIONE_LINK = CreateLink() dal DB destinazione a DB corrente
                ' ============================================================
                
                Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Creazione nuovo link...")
                Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Tabella: '" & sName & "'")
                Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Database sorgente: '" & sDatabasePath & "'")
                
                If CreateLink(sName, sDatabasePath, logFile) Then
                    sStatus = "OK"
                    Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Link CREATO con successo!")
                    Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Stato assegnato: " & sStatus)
                Else
                    sStatus = "NOT_REGISTERED"
                    Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] ERRORE: Impossibile creare il link!")
                    Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Stato assegnato: " & sStatus)
                End If
            Else
                sStatus = "NOT_REGISTERED"
                Call WriteLog(logFile, "[VERIFICA_TABELLA] Tabella '" & sName & "' NON trovata nel database di destinazione.")
                Call WriteLog(logFile, "[VERIFICA_TABELLA] Stato assegnato: " & sStatus)
            End If
            
        End If
        
'======================================================================
' 13_AGGIORNAMENTO_RECORDSET = Update campo Verifica_s
'======================================================================
        Call WriteLog(logFile, "[AGGIORNAMENTO] Salvataggio stato nel campo Verifica_s...")
        
        If rs.Updatable Then
            rs.Edit
            rs!Verifica_s = sStatus
            rs.Update
            Call WriteLog(logFile, "[AGGIORNAMENTO] Recordset aggiornato con stato: '" & sStatus & "' nel campo Verifica_s")
        Else
            Call WriteLog(logFile, "[AGGIORNAMENTO] WARNING: Recordset NON aggiornabile per " & sName)
        End If

'======================================================================
' 14_AGGIORNAMENTO_STATISTICHE = Incremento contatori per stato
'======================================================================
        Select Case sStatus
            Case "OK": nOk = nOk + 1
            Case "MISSING": nMissing = nMissing + 1
            Case "INVALID_PATH": nInvalidPath = nInvalidPath + 1
            Case "NOT_REGISTERED": nNotRegistered = nNotRegistered + 1
        End Select
        
        Call WriteLog(logFile, "[STATISTICHE] Aggiornamento contatori:")
        Call WriteLog(logFile, "            ? OK: " & nOk)
        Call WriteLog(logFile, "            ? MISSING: " & nMissing)
        Call WriteLog(logFile, "            ? INVALID_PATH: " & nInvalidPath)
        Call WriteLog(logFile, "            ? NOT_REGISTERED: " & nNotRegistered)
        Call WriteLog(logFile, "            ? Totale elaborati: " & nTot)
        
        Call WriteLog(logFile, "----------------------------------------")
        Call WriteLog(logFile, "")
        
'======================================================================
' 15_RECORD_SUCCESSIVO = MoveNext
'======================================================================
        rs.MoveNext

    Loop

'======================================================================
' 16_STATISTICHE_FINALI = Report stato sistema su Debug e Log
'======================================================================
    endTime = Now
    
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "STATISTICHE FINALI")
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "Data/Ora fine: " & Format(endTime, "dd/mm/yyyy HH:MM:SS"))
    Call WriteLog(logFile, "Durata totale: " & Format(endTime - startTime, "hh:mm:ss"))
    Call WriteLog(logFile, "")
    Call WriteLog(logFile, "TOTALE RECORD ELABORATI: " & nTot)
    Call WriteLog(logFile, "? OK (link ricostruiti): " & nOk)
    Call WriteLog(logFile, "? MISSING (DB non trovato): " & nMissing)
    Call WriteLog(logFile, "? INVALID_PATH (path errato): " & nInvalidPath)
    Call WriteLog(logFile, "? NOT_REGISTERED (tabella inesistente): " & nNotRegistered)
    Call WriteLog(logFile, "========================================================================")
    
    ' Output anche su Debug
    Debug.Print "========== STATISTICHE FINALI =========="
    Debug.Print "TOTALE RECORD: " & nTot
    Debug.Print "? OK (link ricostruiti): " & nOk
    Debug.Print "? MISSING (DB non trovato): " & nMissing
    Debug.Print "? INVALID_PATH (path errato): " & nInvalidPath
    Debug.Print "? NOT_REGISTERED (tabella inesistente): " & nNotRegistered
    Debug.Print "========================================"

'======================================================================
' 17_CHIUSURA_RISORSE = Cleanup DAO, FSO e file di log
'======================================================================
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set fso = Nothing
    
    Call WriteLog(logFile, "[CHIUSURA] File di log salvato in: " & logFilePath)
    logFile.Close
    Set logFile = Nothing

'======================================================================
' 18_MESSAGGIO_FINALE = MsgBox con esito finale operazione
'======================================================================
    MsgBox "Ricostruzione link completata!" & vbCrLf & vbCrLf & _
           "Totale elaborati: " & nTot & vbCrLf & _
           "? Link OK: " & nOk & vbCrLf & _
           "? DB non trovato (MISSING): " & nMissing & vbCrLf & _
           "? Path invalido (INVALID_PATH): " & nInvalidPath & vbCrLf & _
           "? Tabella non censita (NOT_REGISTERED): " & nNotRegistered & vbCrLf & vbCrLf & _
           "Log salvato in: " & logFilePath, vbInformation

    Exit Sub

'======================================================================
' 99_GESTIONE_ERRORI = Error handler globale con logging
'======================================================================
ErrorHandler:

    Dim errorMsg As String
    errorMsg = "Errore: " & Err.Description & vbCrLf & _
               "Numero: " & Err.Number & vbCrLf & _
               "In tabella: " & sName
    
    If Not logFile Is Nothing Then
        Call WriteLog(logFile, "[ERRORE_GLOBALE] " & errorMsg)
        Call WriteLog(logFile, "[FINE] Operazione terminata con errore")
        logFile.Close
    End If
    
    MsgBox errorMsg, vbCritical

End Sub

'======================================================================
' FUNZIONE DI SCRITTURA LOG
'======================================================================

Private Sub WriteLog(ByVal logFile As Object, ByVal message As String)
    ' Scrive su file di log
    On Error Resume Next
    logFile.WriteLine Now & " - " & message
    On Error GoTo 0
    
    ' Scrive anche su Debug (finestra Immediate)
    Debug.Print Now & " - " & message
End Sub

'======================================================================
' FUNZIONI DI SUPPORTO CON LOGGING
'======================================================================

Private Function QueryExists(ByVal queryName As String) As Boolean
    Dim rs As DAO.Recordset
    
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=5 AND Name='" & queryName & "'")
    On Error GoTo 0
    
    If Not rs Is Nothing Then
        QueryExists = Not rs.EOF
        rs.Close
        Set rs = Nothing
    Else
        QueryExists = False
    End If
End Function

Private Function LinkExistsInMSysObjects(ByVal tableName As String) As Boolean
    Dim rs As DAO.Recordset
    
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Type=6 AND Name='" & tableName & "'")
    On Error GoTo 0
    
    If Not rs Is Nothing Then
        LinkExistsInMSysObjects = Not rs.EOF
        rs.Close
        Set rs = Nothing
    Else
        LinkExistsInMSysObjects = False
    End If
End Function

Private Function TableExistsInDatabase(ByVal databasePath As String, ByVal tableName As String, ByVal logFile As Object) As Boolean
    Dim dbExternal As DAO.Database
    Dim rs As DAO.Recordset
    
    Call WriteLog(logFile, "[APERTURA_DB] Tentativo apertura database: '" & databasePath & "'")
    
    ' Apertura database di destinazione
    On Error Resume Next
    Set dbExternal = OpenDatabase(databasePath)
    
    If Err.Number <> 0 Then
        Call WriteLog(logFile, "[APERTURA_DB] ERRORE: Impossibile aprire il database!")
        Call WriteLog(logFile, "[APERTURA_DB] Errore n. " & Err.Number & ": " & Err.Description)
        TableExistsInDatabase = False
        Exit Function
    End If
    On Error GoTo 0
    
    Call WriteLog(logFile, "[APERTURA_DB] Database aperto con successo.")
    Call WriteLog(logFile, "[VERIFICA_TABELLA] Ricerca tabella '" & tableName & "' in MSysObjects...")
    
    ' Verifica esistenza tabella nel database di destinazione
    ' Type=1 = tabelle locali, Type=4 = tabelle ODBC, Type=6 = link
    Set rs = dbExternal.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE (Type=1 OR Type=4 OR Type=6) AND Name='" & tableName & "'")
    
    TableExistsInDatabase = Not rs.EOF
    
    If Not TableExistsInDatabase Then
        Call WriteLog(logFile, "[VERIFICA_TABELLA] Tabella '" & tableName & "' NON TROVATA")
    Else
        Call WriteLog(logFile, "[VERIFICA_TABELLA] Tabella '" & tableName & "' TROVATA")
    End If
    
    ' 10A_CHIUSURA_DB_DESTINAZIONE = Close database e rilascio risorse
    Call WriteLog(logFile, "[CHIUSURA_DB] Chiusura database di destinazione...")
    rs.Close
    dbExternal.Close
    Set rs = Nothing
    Set dbExternal = Nothing
    Call WriteLog(logFile, "[CHIUSURA_DB] Database chiuso e risorse rilasciate.")
    
End Function

Private Sub DeleteLink(ByVal tableName As String, ByVal logFile As Object)
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Call WriteLog(logFile, "[DELETE_LINK] Eliminazione link '" & tableName & "' da TableDefs...")
    
    On Error Resume Next
    db.TableDefs.Delete tableName
    If Err.Number = 0 Then
        Call WriteLog(logFile, "[DELETE_LINK] Link '" & tableName & "' eliminato con successo.")
    Else
        Call WriteLog(logFile, "[DELETE_LINK] Link non trovato o errore: " & Err.Description)
    End If
    On Error GoTo 0
    
    Set db = Nothing
End Sub

Private Function CreateLink(ByVal tableName As String, ByVal sourceDatabasePath As String, ByVal logFile As Object) As Boolean
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    
    Set db = CurrentDb
    
    Call WriteLog(logFile, "[CREATE_LINK] Creazione TableDef per '" & tableName & "'...")
    
    On Error GoTo CreateLinkError
    
    ' Crea un nuovo TableDef per il link
    Set td = db.CreateTableDef(tableName)
    Call WriteLog(logFile, "[CREATE_LINK] TableDef creato.")
    
    ' Imposta la connessione al database di destinazione
    td.Connect = ";DATABASE=" & sourceDatabasePath
    td.SourceTableName = tableName
    Call WriteLog(logFile, "[CREATE_LINK] Connect string impostata")
    Call WriteLog(logFile, "[CREATE_LINK] SourceTableName impostato: '" & tableName & "'")
    
    ' Aggiungi il link al database corrente
    Call WriteLog(logFile, "[CREATE_LINK] Appendo TableDef al database...")
    db.TableDefs.Append td
    Call WriteLog(logFile, "[CREATE_LINK] TableDef appeso con successo.")
    
    ' Aggiorna il link
    Call WriteLog(logFile, "[CREATE_LINK] Refresh del link...")
    td.RefreshLink
    Call WriteLog(logFile, "[CREATE_LINK] Link aggiornato con successo.")
    
    CreateLink = True
    Call WriteLog(logFile, "[CREATE_LINK] Link creato correttamente per '" & tableName & "'")
    
    Set td = Nothing
    Set db = Nothing
    Exit Function
    
CreateLinkError:
    Call WriteLog(logFile, "[CREATE_LINK] ERRORE: " & Err.Description)
    CreateLink = False
    Set td = Nothing
    Set db = Nothing
End Function

