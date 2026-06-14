Attribute VB_Name = "Msys_TB99_LINK_ Mdl01_RICOSTRUICI_LINK_TABELLE"
Option Compare Database
Option Explicit

'======================================================================
' MODULO     : Msys_TB99_LINK_ Mdl01_RICOSTRUICI_LINK_TABELLE
' PATH       : c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB99_LINK\
' DB         : MSys_TB99_LINK.mdb
' PATH_LOG   : \LOGS
' PROCEDURA  : RICOSTRUICI_LINK_TABELLE_pf
' CODICE     : @SHADOW_SYSTEM_CATALOG
' AUTORE     : System
' DATA       : 2026
' VERSIONE   : 3.0 TB99 SELF-HEALING LINK ENGINE WITH LOGGING
' CODICE_EST : IL codice di estrazione di tutti gli oggetti è --->  Msys_TB99_LINK
'
' FAQ        : @TB99@LINK_(modulo per la ricostruzione delle Tabelle Link di tutti i subprogetti)
'            : @COME@RICOSTRUIRE@_(I COLLEGAMENTI TRA LE TABELLE, @RICOSTRUIRE@LINK)
'
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
' MARKER LOGICI DI FLUSSO (ALLINEATO AL CODICE REALE) +
' NEL CODICE SOTTOETICHETTE NUMERATE ES:  01_01[SET_DB] EC...
'======================================================================
' 01_VARIABILI_PER_LOGGING          = Inizializzazione variabili / DB corrente
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
' 19_CANCELLAZIONE_LINK_ESISTENTI   = Cancella TUTTI i link esistenti (non le tabelle fisiche)
' 99_GESTIONE_ERRORI                = Error handler globale con logging
'======================================================================


'======================================================================
' 01_VARIABILI_PER_LOGGING = Inizializzazione variabili
'======================================================================
Private Const MODULE_PATH As String = "C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB99_LINK\"
Private LOG_FOLDER As String
Private Const LOG_FILE_NAME As String = "TB99_Link_Rebuild_Log.txt"



'======================================================================
' 02_PROCEDURA_PRINCIPALE = RICOSTRUICI_LINK_TABELLE_pf - Avvia la ricostruzione dei link tabelle
' codice ----> @RICOSTRUICI_LINK_TABELLE_pf
'======================================================================
' *** INIZIO ***
Public Function RICOSTRUICI_LINK_TABELLE_pf()

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
' 01_01[SET_DB] = Imposta database corrente
    Set db = CurrentDb
' 01_02[SET_FSO] = Crea oggetto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
' 01_03[TIMER_START] = Avvia contatore temporale
    startTime = Now

'======================================================================
' 01A_CREAZIONE_LOG = Creazione file di log con timestamp
'======================================================================
' 01A_01[SET_LOG_FOLDER] = Imposta percorso cartella LOGS
    LOG_FOLDER = MODULE_PATH & "LOGS\"
    
' 01A_02[CREATE_LOG_FOLDER] = Crea cartella LOG se non esiste
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
    
' 01A_03[BUILD_LOG_PATH] = Costruisce percorso completo file di log
    logFilePath = LOG_FOLDER & Format(startTime, "YYYYMMDD_HHMMSS") & "_" & LOG_FILE_NAME
    
' 01A_04[CREATE_LOG_FILE] = Crea il file di log
    Set logFile = fso.CreateTextFile(logFilePath, True)
    
' 01A_05[WRITE_HEADER] = Scrive intestazione nel log
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
' 02_01[CHECK_QUERY] = Verifica esistenza query Msys_TB99_LINK_Qry99_01_SELECT_Archivio
    Call WriteLog(logFile, "[VERIFICA_QUERY] Controllo esistenza query Msys_TB99_LINK_Qry99_01_SELECT_Archivio...")
    
    If Not QueryExists("Msys_TB99_LINK_Qry99_01_SELECT_Archivio") Then
        Call WriteLog(logFile, "[ERRORE] Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio NON TROVATA!")
        Call WriteLog(logFile, "[FINE] Operazione terminata con errore")
        logFile.Close
        Set logFile = Nothing
        MsgBox "Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio non trovata", vbCritical
        Exit Function
    End If
    
    Call WriteLog(logFile, "[VERIFICA_QUERY] Query trovata con successo.")
    Call WriteLog(logFile, "")

'======================================================================
' 03_CONFERMA_UTENTE = Conferma operazione da parte dell'utente
'======================================================================
' 03_01[ASK_CONFIRM] = Richiede conferma all'utente
    Call WriteLog(logFile, "[UTENTE] Richiesta conferma operazione...")
    
    If MsgBox("Avvio ricostruzione link tabelle da TB99?", vbYesNo) = vbNo Then
        Call WriteLog(logFile, "[UTENTE] Operazione ANNULLATA dall'utente.")
        Call WriteLog(logFile, "[FINE] Log chiuso.")
        logFile.Close
        Set logFile = Nothing
        Exit Function
    End If
    
    Call WriteLog(logFile, "[UTENTE] Operazione CONFERMATA.")
    Call WriteLog(logFile, "")

'======================================================================
' 04_APERTURA_RECORDSET_QUERY = Apertura query Msys_TB99_LINK_Qry99_01_SELECT_Archivio
'======================================================================
' 04_01[BUILD_SQL] = Costruisce query SQL
    sSQL = "SELECT * FROM Msys_TB99_LINK_Qry99_01_SELECT_Archivio"
' 04_02[OPEN_RECORDSET] = Apre recordset
    Set rs = db.OpenRecordset(sSQL, dbOpenDynaset)
    
    Call WriteLog(logFile, "[APERTURA_QUERY] Recordset aperto. Record trovati: " & rs.RecordCount)
    Call WriteLog(logFile, "")

'======================================================================
' 05_INIZIALIZZAZIONE_STATISTICHE = Reset contatori statistiche
'======================================================================
' 05_01[RESET_COUNTERS] = Inizializza contatori a zero
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
    
' 06_01[LOOP_START] = Inizio ciclo Do While
    Do While Not rs.EOF
        nTot = nTot + 1
        
        Call WriteLog(logFile, "----------------------------------------")
        Call WriteLog(logFile, "RECORD N. " & nTot)
        Call WriteLog(logFile, "----------------------------------------")

'======================================================================
' 07_LETTURA_RECORD = Lettura Name, Database, Type
'======================================================================
' 07_01[READ_NAME] = Legge campo Name
        sName = Nz(rs!Name, "")
' 07_02[READ_DATABASE] = Legge campo Database
        sDatabasePath = Nz(rs!Database, "")
' 07_03[READ_TYPE] = Legge campo Type
        lType = Nz(rs!Type, 0)

        Call WriteLog(logFile, "[LETTURA_RECORD] Nome tabella: '" & sName & "'")
        Call WriteLog(logFile, "[LETTURA_RECORD] Database path: '" & sDatabasePath & "'")
        Call WriteLog(logFile, "[LETTURA_RECORD] Type: " & lType)
        
' 07_04[INIT_STATUS] = Inizializza stato
        sStatus = "NOT_CHECKED"

'======================================================================
' 08_VALIDAZIONE_PATH = Verifica Type=6 e path valido
'======================================================================
        Call WriteLog(logFile, "[VALIDAZIONE] Avvio validazione...")
        
' 08_01[CHECK_TYPE] = Verifica che sia un link table (Type = 6)
        If lType <> 6 Then
            sStatus = "NOT_REGISTERED"
            Call WriteLog(logFile, "[VALIDAZIONE] Type = " & lType & " (diverso da 6) -> NON è un link table")
            Call WriteLog(logFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
' 08_02[CHECK_PATH_FORMAT] = Verifica path strutturalmente valido
        ElseIf Len(Trim(sDatabasePath)) = 0 Or InStr(1, sDatabasePath, ":\") = 0 Then
            sStatus = "INVALID_PATH"
            Call WriteLog(logFile, "[VALIDAZIONE] Path strutturalmente INVALIDO")
            Call WriteLog(logFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
' 08_03[CHECK_FILE_EXISTS] = Verifica esistenza file fisico del database
        ElseIf Not fso.FileExists(sDatabasePath) Then
            sStatus = "MISSING"
            Call WriteLog(logFile, "[VALIDAZIONE] File database NON TROVATO sul filesystem")
            Call WriteLog(logFile, "[VALIDAZIONE] Path cercato: '" & sDatabasePath & "'")
            Call WriteLog(logFile, "[VALIDAZIONE] Stato assegnato: " & sStatus)
            
        Else
            Call WriteLog(logFile, "[VALIDAZIONE] Path valido. File database ESISTE.")
            Call WriteLog(logFile, "[VALIDAZIONE] Inizio verifica lato SORGENTE...")
            
'======================================================================
' 09_APERTURA_DB_DESTINAZIONE = OpenDatabase() sul campo Database
' 10_VERIFICA_TABELLA_SORGENTE = Controllo esistenza tabella nel DB destinazione
' 10A_CHIUSURA_DB_DESTINAZIONE = Close database e rilascio risorse
'======================================================================
            
' 09_01[OPEN_EXTERNAL_DB] = Apre il database di destinazione
' 10_01[CHECK_TABLE_EXISTS] = Verifica esistenza tabella
' 10A_01[CLOSE_EXTERNAL_DB] = Chiude database e rilascia risorse
            tableExistsInSource = TableExistsInDatabase(sDatabasePath, sName, logFile)
            
' 10_02[EVALUATE_RESULT] = Valuta risultato verifica
            If tableExistsInSource Then
                Call WriteLog(logFile, "[VERIFICA_TABELLA] Tabella '" & sName & "' TROVATA nel database di destinazione.")
                
'======================================================================
' 11_ELIMINAZIONE_LINK_PRECEDENTE = DeleteLink() se esiste
'======================================================================
                
' 11_01[CHECK_LINK_EXISTS] = Controlla se link esiste già
                If LinkExistsInMSysObjects(sName) Then
                    Call WriteLog(logFile, "[ELIMINAZIONE_LINK] Link esistente TROVATO in MSysObjects.")
                    Call WriteLog(logFile, "[ELIMINAZIONE_LINK] Eliminazione in corso...")
' 11_02[DELETE_LINK] = Elimina link esistente
                    DeleteLink sName, logFile
                Else
                    Call WriteLog(logFile, "[ELIMINAZIONE_LINK] Link esistente NON trovato in MSysObjects.")
                End If
                
'======================================================================
' 12_RICOSTRUZIONE_LINK = CreateLink() dal DB destinazione a DB corrente
'======================================================================
                
                Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Creazione nuovo link...")
                Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Tabella: '" & sName & "'")
                Call WriteLog(logFile, "[RICOSTRUZIONE_LINK] Database sorgente: '" & sDatabasePath & "'")
                
' 12_01[CREATE_NEW_LINK] = Crea nuovo link
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
        
' 13_01[CHECK_UPDATABLE] = Verifica se recordset è aggiornabile
        If rs.Updatable Then
' 13_02[EDIT_RECORD] = Entra in modalità modifica
            rs.Edit
' 13_03[UPDATE_STATUS] = Aggiorna campo Verifica_s
            rs!Verifica_s = sStatus

' 13_04[UPDATE_SUBPROJECT] = Aggiorna campo COD_SUB_PROGETTO_c con primi 9 caratteri di Name
            If Len(sName) >= 9 Then
                rs!COD_SUB_PROGETTO_c = Left(sName, 9)
                Call WriteLog(logFile, "[AGGIORNAMENTO] COD_SUB_PROGETTO_c impostato a: '" & Left(sName, 9) & "'")
            Else
                rs!COD_SUB_PROGETTO_c = sName
                Call WriteLog(logFile, "[AGGIORNAMENTO] COD_SUB_PROGETTO_c impostato a (nome completo <9 char): '" & sName & "'")
            End If
            
' 13_05[SAVE_RECORD] = Salva modifiche
            rs.Update
            Call WriteLog(logFile, "[AGGIORNAMENTO] Recordset aggiornato con stato: '" & sStatus & "' nel campo Verifica_s")
        Else
            Call WriteLog(logFile, "[AGGIORNAMENTO] WARNING: Recordset NON aggiornabile per " & sName)
        End If

'======================================================================
' 14_AGGIORNAMENTO_STATISTICHE = Incremento contatori per stato
'======================================================================
' 14_01[INCREMENT_OK] = Incrementa contatore OK
' 14_02[INCREMENT_MISSING] = Incrementa contatore MISSING
' 14_03[INCREMENT_INVALID] = Incrementa contatore INVALID_PATH
' 14_04[INCREMENT_NOTREG] = Incrementa contatore NOT_REGISTERED
        Select Case sStatus
            Case "OK": nOk = nOk + 1
            Case "MISSING": nMissing = nMissing + 1
            Case "INVALID_PATH": nInvalidPath = nInvalidPath + 1
            Case "NOT_REGISTERED": nNotRegistered = nNotRegistered + 1
        End Select
        
        Call WriteLog(logFile, "[STATISTICHE] Aggiornamento contatori:")
        Call WriteLog(logFile, "            OK: " & nOk)
        Call WriteLog(logFile, "            MISSING: " & nMissing)
        Call WriteLog(logFile, "            INVALID_PATH: " & nInvalidPath)
        Call WriteLog(logFile, "            NOT_REGISTERED: " & nNotRegistered)
        Call WriteLog(logFile, "            Totale elaborati: " & nTot)
        
        Call WriteLog(logFile, "----------------------------------------")
        Call WriteLog(logFile, "")
        
'======================================================================
' 15_RECORD_SUCCESSIVO = MoveNext
'======================================================================
' 15_01[NEXT_RECORD] = Passa al record successivo
        rs.MoveNext

    Loop
' 06_02[LOOP_END] = Fine ciclo Do While

'======================================================================
' 16_STATISTICHE_FINALI = Report stato sistema su Debug e Log
'======================================================================
' 16_01[TIMER_END] = Rileva ora fine elaborazione
    endTime = Now
    
' 16_02[WRITE_FINAL_STATS] = Scrive statistiche finali su log
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "STATISTICHE FINALI")
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "Data/Ora fine: " & Format(endTime, "dd/mm/yyyy HH:MM:SS"))
    Call WriteLog(logFile, "Durata totale: " & Format(endTime - startTime, "hh:mm:ss"))
    Call WriteLog(logFile, "")
    Call WriteLog(logFile, "TOTALE RECORD ELABORATI: " & nTot)
    Call WriteLog(logFile, "OK (link ricostruiti): " & nOk)
    Call WriteLog(logFile, "MISSING (DB non trovato): " & nMissing)
    Call WriteLog(logFile, "INVALID_PATH (path errato): " & nInvalidPath)
    Call WriteLog(logFile, "NOT_REGISTERED (tabella inesistente): " & nNotRegistered)
    Call WriteLog(logFile, "========================================================================")
    
' 16_03[DEBUG_PRINT] = Stampa statistiche su finestra Immediate
    Debug.Print "========== STATISTICHE FINALI =========="
    Debug.Print "TOTALE RECORD: " & nTot
    Debug.Print "OK (link ricostruiti): " & nOk
    Debug.Print "MISSING (DB non trovato): " & nMissing
    Debug.Print "INVALID_PATH (path errato): " & nInvalidPath
    Debug.Print "NOT_REGISTERED (tabella inesistente): " & nNotRegistered
    Debug.Print "========================================"

'======================================================================
' 17_CHIUSURA_RISORSE = Cleanup DAO, FSO e file di log
'======================================================================
' 17_01[CLOSE_RECORDSET] = Chiude recordset
    rs.Close
' 17_02[RELEASE_RS] = Rilascia recordset
    Set rs = Nothing
' 17_03[RELEASE_DB] = Rilascia database
    Set db = Nothing
' 17_04[RELEASE_FSO] = Rilascia FileSystemObject
    Set fso = Nothing
    
    Call WriteLog(logFile, "[CHIUSURA] File di log salvato in: " & logFilePath)
' 17_05[CLOSE_LOG] = Chiude file di log
    logFile.Close
' 17_06[RELEASE_LOG] = Rilascia oggetto log
    Set logFile = Nothing

'======================================================================
' 18_MESSAGGIO_FINALE = MsgBox con esito finale operazione
'======================================================================
' 18_01[SHOW_RESULT] = Mostra messaggio riepilogativo all'utente
    MsgBox "Ricostruzione link completata!" & vbCrLf & vbCrLf & _
           "Totale elaborati: " & nTot & vbCrLf & _
           "Link OK: " & nOk & vbCrLf & _
           "DB non trovato (MISSING): " & nMissing & vbCrLf & _
           "Path invalido (INVALID_PATH): " & nInvalidPath & vbCrLf & _
           "Tabella non censita (NOT_REGISTERED): " & nNotRegistered & vbCrLf & vbCrLf & _
           "Log salvato in: " & logFilePath, vbInformation

    Exit Function

'======================================================================
' 99_GESTIONE_ERRORI = Error handler globale con logging
'======================================================================
ErrorHandler:
' 99_01[BUILD_ERROR_MSG] = Costruisce messaggio errore
    Dim errorMsg As String
    errorMsg = "Errore: " & Err.Description & vbCrLf & _
               "Numero: " & Err.Number & vbCrLf & _
               "In tabella: " & sName
    
' 99_02[LOG_ERROR] = Scrive errore nel log
    If Not logFile Is Nothing Then
        Call WriteLog(logFile, "[ERRORE_GLOBALE] " & errorMsg)
        Call WriteLog(logFile, "[FINE] Operazione terminata con errore")
        logFile.Close
    End If
    
' 99_03[SHOW_ERROR] = Mostra errore all'utente
    MsgBox errorMsg, vbCritical

End Function
' *** FINE ***


'======================================================================
' 22_FUNZIONE_LOG = WriteLog - Scrive messaggi su file di log e debug
'======================================================================
' *** INIZIO ***
Private Sub WriteLog(ByVal logFile As Object, ByVal message As String)
' 22_01[WRITE_TO_FILE] = Scrive su file di log
    On Error Resume Next
    logFile.WriteLine Now & " - " & message
    On Error GoTo 0
    
' 22_02[WRITE_TO_DEBUG] = Scrive su finestra Immediate
    Debug.Print Now & " - " & message
End Sub
' *** FINE ***


'======================================================================
' 23_VERIFICA_QUERY = QueryExists - Controlla esistenza query in MSysObjects
'======================================================================
' *** INIZIO ***
Private Function QueryExists(ByVal queryName As String) As Boolean
    Dim rs As DAO.Recordset
    
' 23_01[OPEN_MSYSOBJECTS] = Cerca query in MSysObjects (Type=5)
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=5 AND Name='" & queryName & "'")
    On Error GoTo 0
    
' 23_02[CHECK_RECORD] = Verifica se query esiste
    If Not rs Is Nothing Then
        QueryExists = Not rs.EOF
' 23_03[CLOSE_RS] = Chiude recordset
        rs.Close
        Set rs = Nothing
    Else
        QueryExists = False
    End If
End Function
' *** FINE ***


'======================================================================
' 24_VERIFICA_LINK = LinkExistsInMSysObjects - Controlla esistenza link in MSysObjects
'======================================================================
' *** INIZIO ***
Private Function LinkExistsInMSysObjects(ByVal tableName As String) As Boolean
    Dim rs As DAO.Recordset
    
' 24_01[OPEN_MSYSOBJECTS] = Cerca link in MSysObjects (Type=6)
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Type=6 AND Name='" & tableName & "'")
    On Error GoTo 0
    
' 24_02[CHECK_RECORD] = Verifica se link esiste
    If Not rs Is Nothing Then
        LinkExistsInMSysObjects = Not rs.EOF
' 24_03[CLOSE_RS] = Chiude recordset
        rs.Close
        Set rs = Nothing
    Else
        LinkExistsInMSysObjects = False
    End If
End Function
' *** FINE ***


'======================================================================
' 25_VERIFICA_TABELLA = TableExistsInDatabase - Verifica esistenza tabella in DB esterno
'======================================================================
' *** INIZIO ***
Private Function TableExistsInDatabase(ByVal databasePath As String, ByVal tableName As String, ByVal logFile As Object) As Boolean
    Dim dbExternal As DAO.Database
    Dim rs As DAO.Recordset
    
    Call WriteLog(logFile, "[APERTURA_DB] Tentativo apertura database: '" & databasePath & "'")
    
' 25_01[OPEN_EXTERNAL_DB] = Tenta apertura database esterno
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
    
' 25_02[CHECK_TABLE_IN_EXTERNAL] = Cerca tabella nel DB esterno (Type=1,4,6)
    Set rs = dbExternal.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE (Type=1 OR Type=4 OR Type=6) AND Name='" & tableName & "'")
    
    TableExistsInDatabase = Not rs.EOF
    
    If Not TableExistsInDatabase Then
        Call WriteLog(logFile, "[VERIFICA_TABELLA] Tabella '" & tableName & "' NON TROVATA")
    Else
        Call WriteLog(logFile, "[VERIFICA_TABELLA] Tabella '" & tableName & "' TROVATA")
    End If
    
'======================================================================
' 10A_CHIUSURA_DB_DESTINAZIONE = Close database e rilascio risorse
'======================================================================
' 25_03[CLOSE_EXTERNAL_RS] = Chiude recordset esterno
    rs.Close
' 25_04[CLOSE_EXTERNAL_DB] = Chiude database esterno
    dbExternal.Close
' 25_05[RELEASE_EXTERNAL_OBJECTS] = Rilascia oggetti
    Set rs = Nothing
    Set dbExternal = Nothing
    Call WriteLog(logFile, "[CHIUSURA_DB] Database chiuso e risorse rilasciate.")
    
End Function
' *** FINE ***


'======================================================================
' 27_ELIMINAZIONE_LINK = DeleteLink - Elimina link tabella da TableDefs
'======================================================================
' *** INIZIO ***
Private Sub DeleteLink(ByVal tableName As String, ByVal logFile As Object)
    Dim db As DAO.Database
    
' 27_01[SET_CURRENT_DB] = Imposta database corrente
    Set db = CurrentDb
    
    Call WriteLog(logFile, "[DELETE_LINK] Eliminazione link '" & tableName & "' da TableDefs...")
    
' 27_02[DELETE_TABLEDEF] = Elimina TableDef dal database
    On Error Resume Next
    db.TableDefs.Delete tableName
    If Err.Number = 0 Then
        Call WriteLog(logFile, "[DELETE_LINK] Link '" & tableName & "' eliminato con successo.")
    Else
        Call WriteLog(logFile, "[DELETE_LINK] Link non trovato o errore: " & Err.Description)
    End If
    On Error GoTo 0
    
' 27_03[RELEASE_DB] = Rilascia oggetto database
    Set db = Nothing
End Sub
' *** FINE ***


'======================================================================
' 28_CREAZIONE_LINK = CreateLink - Crea nuovo link tabella verso database esterno
'======================================================================
' *** INIZIO ***
Private Function CreateLink(ByVal tableName As String, ByVal sourceDatabasePath As String, ByVal logFile As Object) As Boolean
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    
' 28_01[SET_CURRENT_DB] = Imposta database corrente
    Set db = CurrentDb
    
    Call WriteLog(logFile, "[CREATE_LINK] Creazione TableDef per '" & tableName & "'...")
    
    On Error GoTo CreateLinkError
    
' 28_02[CREATE_TABLEDEF] = Crea nuovo TableDef per il link
    Set td = db.CreateTableDef(tableName)
    Call WriteLog(logFile, "[CREATE_LINK] TableDef creato.")
    
' 28_03[SET_CONNECT_STRING] = Imposta stringa di connessione
    td.Connect = ";DATABASE=" & sourceDatabasePath
' 28_04[SET_SOURCE_TABLE] = Imposta nome tabella sorgente
    td.SourceTableName = tableName
    Call WriteLog(logFile, "[CREATE_LINK] Connect string impostata")
    Call WriteLog(logFile, "[CREATE_LINK] SourceTableName impostato: '" & tableName & "'")
    
' 28_05[APPEND_TABLEDEF] = Aggiunge TableDef al database corrente
    Call WriteLog(logFile, "[CREATE_LINK] Appendo TableDef al database...")
    db.TableDefs.Append td
    Call WriteLog(logFile, "[CREATE_LINK] TableDef appeso con successo.")
    
' 28_06[REFRESH_LINK] = Aggiorna il link
    Call WriteLog(logFile, "[CREATE_LINK] Refresh del link...")
    td.RefreshLink
    Call WriteLog(logFile, "[CREATE_LINK] Link aggiornato con successo.")
    
    CreateLink = True
    Call WriteLog(logFile, "[CREATE_LINK] Link creato correttamente per '" & tableName & "'")
    
' 28_07[RELEASE_OBJECTS] = Rilascia oggetti
    Set td = Nothing
    Set db = Nothing
    Exit Function
    
CreateLinkError:
' 28_99[ERROR_HANDLER] = Gestione errore creazione link
    Call WriteLog(logFile, "[CREATE_LINK] ERRORE: " & Err.Description)
    CreateLink = False
    Set td = Nothing
    Set db = Nothing
End Function
' *** FINE ***

'======================================================================
' 19_CANCELLAZIONE_LINK_ESISTENTI = Cancella TUTTI i link esistenti (non le tabelle fisiche)
'======================================================================
' *** INIZIO ***
Public Function CancellaTuttiLinkEsistenti() As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fso As Object
    Dim logFile As Object
    Dim logFilePath As String
    Dim sSQL As String
    Dim sName As String
    Dim nTot As Long
    Dim nCancellati As Long
    Dim nErrori As Long
    Dim startTime As Date
    Dim endTime As Date
    
    On Error GoTo ErrorHandler
    
' 19_01[INIT_VARS] = Inizializzazione variabili
    Set db = CurrentDb
    Set fso = CreateObject("Scripting.FileSystemObject")
    startTime = Now
    
' 19_02[CREATE_LOG] = Crea file di log per cancellazione
    LOG_FOLDER = MODULE_PATH & "LOGS\"
    
    If Not fso.FolderExists(LOG_FOLDER) Then
        On Error Resume Next
        fso.CreateFolder LOG_FOLDER
        If Err.Number <> 0 Then
            LOG_FOLDER = "C:\TEMP\"
            fso.CreateFolder LOG_FOLDER
        End If
        On Error GoTo 0
    End If
    
    logFilePath = LOG_FOLDER & Format(startTime, "YYYYMMDD_HHMMSS") & "_DELETE_LINKS_" & LOG_FILE_NAME
    Set logFile = fso.CreateTextFile(logFilePath, True)
    
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "INIZIO CANCELLAZIONE TUTTI I LINK ESISTENTI")
    Call WriteLog(logFile, "Data/Ora avvio: " & Format(startTime, "dd/mm/yyyy HH:MM:SS"))
    Call WriteLog(logFile, "Database corrente: " & CurrentDb.Name)
    Call WriteLog(logFile, "File di log: " & logFilePath)
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "")
    
' 19_03[VERIFY_QUERY] = Verifica esistenza query TB99
    Call WriteLog(logFile, "[VERIFICA_QUERY] Controllo esistenza query Msys_TB99_LINK_Qry99_01_SELECT_Archivio...")
    
    If Not QueryExists("Msys_TB99_LINK_Qry99_01_SELECT_Archivio") Then
        Call WriteLog(logFile, "[ERRORE] Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio NON TROVATA!")
        Call WriteLog(logFile, "[FINE] Operazione terminata con errore")
        logFile.Close
        Set logFile = Nothing
        MsgBox "Query Msys_TB99_LINK_Qry99_01_SELECT_Archivio non trovata", vbCritical
        CancellaTuttiLinkEsistenti = False
        Exit Function
    End If
    
    Call WriteLog(logFile, "[VERIFICA_QUERY] Query trovata con successo.")
    Call WriteLog(logFile, "")
    
' 19_04[ASK_CONFIRM] = Richiede conferma all'utente
    Call WriteLog(logFile, "[UTENTE] Richiesta conferma cancellazione TUTTI i link...")
    
    If MsgBox("ATTENZIONE! Verranno cancellati TUTTI i link tabelle presenti nella query TB99." & vbCrLf & vbCrLf & _
              "Questa operazione NON cancella le tabelle fisiche, solo i collegamenti." & vbCrLf & vbCrLf & _
              "Procedere?", vbYesNo + vbExclamation) = vbNo Then
        Call WriteLog(logFile, "[UTENTE] Operazione ANNULLATA dall'utente.")
        Call WriteLog(logFile, "[FINE] Log chiuso.")
        logFile.Close
        Set logFile = Nothing
        CancellaTuttiLinkEsistenti = False
        Exit Function
    End If
    
    Call WriteLog(logFile, "[UTENTE] Operazione CONFERMATA.")
    Call WriteLog(logFile, "")
    
' 19_05[OPEN_RECORDSET] = Apertura recordset query TB99
    sSQL = "SELECT * FROM Msys_TB99_LINK_Qry99_01_SELECT_Archivio"
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
    
    Call WriteLog(logFile, "[APERTURA_QUERY] Recordset aperto. Record trovati: " & rs.RecordCount)
    Call WriteLog(logFile, "")
    
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "INIZIO CANCELLAZIONE LINK")
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "")
    
' 19_06[LOOP_START] = Loop su tutti i record della query
    nTot = 0
    nCancellati = 0
    nErrori = 0
    
    Do While Not rs.EOF
        nTot = nTot + 1
        sName = Nz(rs!Name, "")
        
        Call WriteLog(logFile, "----------------------------------------")
        Call WriteLog(logFile, "RECORD N. " & nTot & " - Tabella: '" & sName & "'")
        Call WriteLog(logFile, "----------------------------------------")
        
' 19_07[CHECK_IF_LINK] = Verifica se è effettivamente un link (Type=6)
        If Nz(rs!Type, 0) = 6 Then
            
' 19_08[CHECK_LINK_EXISTS] = Verifica se il link esiste in MSysObjects
            If LinkExistsInMSysObjects(sName) Then
                Call WriteLog(logFile, "[CANCELLAZIONE] Link trovato, eliminazione in corso...")
                
' 19_09[DELETE_LINK] = Elimina il link
                On Error Resume Next
                db.TableDefs.Delete sName
                If Err.Number = 0 Then
                    nCancellati = nCancellati + 1
                    Call WriteLog(logFile, "[CANCELLAZIONE] Link '" & sName & "' eliminato con SUCCESSO.")
                Else
                    nErrori = nErrori + 1
                    Call WriteLog(logFile, "[CANCELLAZIONE] ERRORE eliminazione link '" & sName & "': " & Err.Description)
                End If
                On Error GoTo 0
            Else
                Call WriteLog(logFile, "[CANCELLAZIONE] Link '" & sName & "' NON trovato in MSysObjects (già assente).")
            End If
        Else
            Call WriteLog(logFile, "[CANCELLAZIONE] Type = " & Nz(rs!Type, 0) & " -> NON è un link table, saltato.")
        End If
        
        Call WriteLog(logFile, "----------------------------------------")
        Call WriteLog(logFile, "")
        
' 19_10[NEXT_RECORD] = Passa al record successivo
        rs.MoveNext
    Loop
' 19_11[LOOP_END] = Fine ciclo
    
' 19_12[STATISTICHE_FINALI] = Report statistiche cancellazione
    endTime = Now
    
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "STATISTICHE FINALI CANCELLAZIONE")
    Call WriteLog(logFile, "========================================================================")
    Call WriteLog(logFile, "Data/Ora fine: " & Format(endTime, "dd/mm/yyyy HH:MM:SS"))
    Call WriteLog(logFile, "Durata totale: " & Format(endTime - startTime, "hh:mm:ss"))
    Call WriteLog(logFile, "")
    Call WriteLog(logFile, "TOTALE RECORD ESAMINATI: " & nTot)
    Call WriteLog(logFile, "LINK CANCELLATI CON SUCCESSO: " & nCancellati)
    Call WriteLog(logFile, "ERRORI DURANTE CANCELLAZIONE: " & nErrori)
    Call WriteLog(logFile, "========================================================================")
    
    Debug.Print "========== STATISTICHE CANCELLAZIONE LINK =========="
    Debug.Print "TOTALE RECORD: " & nTot
    Debug.Print "LINK CANCELLATI: " & nCancellati
    Debug.Print "ERRORI: " & nErrori
    Debug.Print "===================================================="
    
' 19_13[CLOSE_RESOURCES] = Chiusura risorse
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set fso = Nothing
    
    Call WriteLog(logFile, "[CHIUSURA] File di log salvato in: " & logFilePath)
    logFile.Close
    Set logFile = Nothing
    
' 19_14[MESSAGGIO_FINALE] = MsgBox con esito finale
    MsgBox "Cancellazione link completata!" & vbCrLf & vbCrLf & _
           "Totale record esaminati: " & nTot & vbCrLf & _
           "Link cancellati con successo: " & nCancellati & vbCrLf & _
           "Errori: " & nErrori & vbCrLf & vbCrLf & _
           "Log salvato in: " & logFilePath, vbInformation
    
    CancellaTuttiLinkEsistenti = True
    Exit Function
    
ErrorHandler:
' 19_99[ERROR_HANDLER] = Gestione errori
    Dim errorMsg As String
    errorMsg = "Errore in CancellaTuttiLinkEsistenti: " & Err.Description & vbCrLf & _
               "Numero: " & Err.Number
    
    If Not logFile Is Nothing Then
        Call WriteLog(logFile, "[ERRORE_GLOBALE] " & errorMsg)
        Call WriteLog(logFile, "[FINE] Operazione terminata con errore")
        logFile.Close
    End If
    
    MsgBox errorMsg, vbCritical
    CancellaTuttiLinkEsistenti = False
End Function
' *** FINE ***



