Attribute VB_Name = "ESPORTA_Mdl01_TUTTI_GLI_OGGETTI_IN_FILE_HTML"
'//    MODULO ---> ESPORTA_Mdl01_TUTTI_GLI_OGGETTI_IN_FILE_HTML
'//****************************************************************************//'
'// ATTIVITA : esporto i seguenti oggetti in formato HTML
'//      01) @esporta@query@HTML_(la procedura ora salva gli oggedtti query in formato html)
'//          @esportare@query_(in un file .log ed in html)
'//          @esportare@Moduli_(in un file .log ed in html)
'//          @esportare@Report_(in un file .log ed in html)
'//
'//   apri db: START "APRI MODULI" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\MSys_MODULI.mdb"
'//   apri db: START "APRI PATH" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\"
'//
'//   faq:
'//      come esporto le query in file HTML
'//      come esportare TUTTI GLI OGGETTI IN @HTML
'//
'//*******************************************************************************************************//



Option Compare Database
Option Explicit


Public Function SalvaTutti_GliOggetti_InHTML()
    Call EsportaTutteLeQueryInHTML
    Call EsportaTuttiIModuliInHTML
    Call EsportaTutteLeFormInHTML
    Call EsportaStrutturaTutteLeTabelleInHTML
End Function




'//======================================================================================================//
'//                         @ESPORTA LE QUERY IN FORMATO HTML                      *** INIZIO ***
'//======================================================================================================//

Sub EsportaTutteLeQueryInHTML()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim percorso As String
    Dim logFile As Integer
    Dim htmlFile As Integer
    Dim htmlString As String
    Dim logPercorso As String
    Dim htmlPercorso As String
    Dim contatoreSalvati As Long ' Contatore query salvate
    
    ' Recupera la cartella corrente del database
    percorso = CurrentProject.Path & "\..\OBJECT\QUERY\"
    logPercorso = percorso & "LOG_EsportaTutteLeQueryInHTML.txt"
    
    ' Crea la cartella se non esiste
    If Dir(percorso, vbDirectory) = "" Then MkDir percorso
    
    ' Apre il file di log in modalità append per tracciare l'esportazione
    logFile = FreeFile
    Open logPercorso For Append As logFile
    Print #logFile, "=== Inizio esportazione SQL in HTML: " & Now() & " ==="
    
    ' Apre il database corrente
    Set db = CurrentDb
    
    contatoreSalvati = 0 ' Inizializza contatore
    
    ' Cicla tutte le query nel database
    For Each qdf In db.QueryDefs
        ' Esclude query di sistema e temporanee (nome che inizia con ~)
        If Left(qdf.Name, 1) <> "~" Then
            On Error GoTo GestioneErrore
            
            ' Costruisce la stringa HTML contenente il nome e il codice SQL formattato
            htmlString = "<!DOCTYPE html>" & vbCrLf & _
                         "<html lang=""it"">" & vbCrLf & _
                         "<head>" & vbCrLf & _
                         "  <meta charset=""UTF-8"">" & vbCrLf & _
                         "  <title>Query: " & qdf.Name & "</title>" & vbCrLf & _
                         "  <style>" & vbCrLf & _
                         "    body { font-family: Consolas, monospace; background:#f0f0f0; padding:20px; }" & vbCrLf & _
                         "    h1 { color:#333; }" & vbCrLf & _
                         "    pre { background:#222; color:#eee; padding:15px; border-radius:5px; white-space: pre-wrap; }" & vbCrLf & _
                         "  </style>" & vbCrLf & _
                         "</head>" & vbCrLf & _
                         "<body>" & vbCrLf & _
                         "  <h1>Query: " & qdf.Name & "</h1>" & vbCrLf & _
                         "  <pre>" & Replace(qdf.sql, "<", "&lt;") & "</pre>" & vbCrLf & _
                         "</body>" & vbCrLf & _
                         "</html>"
            
            ' Salva il file HTML con nome della query nella cartella del DB
            htmlPercorso = percorso & qdf.Name & ".html"
            htmlFile = FreeFile
            Open htmlPercorso For Output As htmlFile
            Print #htmlFile, htmlString
            Close htmlFile
            
            contatoreSalvati = contatoreSalvati + 1 ' Incrementa contatore
            
            ' Scrive nel log il successo dell'esportazione
            Print #logFile, "SUCCESSO - Salvata query '" & qdf.Name & "' in HTML"
            
            On Error GoTo 0  ' reset error handling per il prossimo ciclo
        End If
    Next qdf
    
    ' Scrive la statistica nel log
    Print #logFile, "=== Statistica: query salvate con successo = " & contatoreSalvati & " ==="
    
    ' Segnala la fine del processo di esportazione nel log
    Print #logFile, "=== Fine esportazione SQL in HTML: " & Now() & " ==="
    Close logFile
    
    ' Pulisce variabili per evitare perdite di memoria
    Set qdf = Nothing
    Set db = Nothing
    
    MsgBox "Esportazione SQL in HTML completata!" & vbCrLf & _
           "Query salvate con successo: " & contatoreSalvati, vbInformation
    Exit Sub
    
GestioneErrore:
    Print #logFile, "ERRORE - Durante esportazione query '" & qdf.Name & "': " & Err.Description
    Err.Clear
    Resume Next
End Sub

'//======================================================================================================//
'//                         @ESPORTA LE QUERY IN FORMATO HTML                      *** FINE ***
'//======================================================================================================//



'//======================================================================================================//
'//                         @ESPORTA I MODULI IN FORMATO HTML                      *** INIZIO ***
'//======================================================================================================//

Sub EsportaTuttiIModuliInHTML()
    Dim vbComp As Object ' VBIDE.VBComponent
    Dim codice As String
    Dim i As Long
    Dim percorso As String
    Dim logPercorso As String
    Dim htmlPercorso As String
    Dim logFile As Integer
    Dim htmlFile As Integer
    Dim htmlString As String
    Dim contatoreSalvati As Long ' Contatore moduli salvati
    
    ' Imposta la cartella di esportazione
    percorso = CurrentProject.Path & "\..\OBJECT\MODULI\"
    logPercorso = percorso & "LOG_EsportaTuttiIModuliInHTML.txt"
    
    ' Crea la cartella se non esiste
    If Dir(percorso, vbDirectory) = "" Then MkDir percorso
    
    ' Apre file di log
    logFile = FreeFile
    Open logPercorso For Append As logFile
    Print #logFile, "=== Inizio esportazione moduli in HTML: " & Now() & " ==="
    
    contatoreSalvati = 0 ' Inizializza contatore
    
    ' Scorre tutti i componenti del progetto VBA
    For Each vbComp In Application.VBE.VBProjects(1).VBComponents
        On Error GoTo GestioneErrore
        
        ' Recupera il codice del modulo riga per riga
        codice = ""
        For i = 1 To vbComp.CodeModule.CountOfLines
            codice = codice & vbComp.CodeModule.Lines(i, 1) & vbCrLf
        Next i
        
        ' Costruisce il contenuto HTML
        htmlString = "<!DOCTYPE html>" & vbCrLf & _
                     "<html lang=""it"">" & vbCrLf & _
                     "<head>" & vbCrLf & _
                     "  <meta charset=""UTF-8"">" & vbCrLf & _
                     "  <title>Modulo: " & vbComp.Name & "</title>" & vbCrLf & _
                     "  <style>" & vbCrLf & _
                     "    body { font-family: Consolas, monospace; background:#f0f0f0; padding:20px; }" & vbCrLf & _
                     "    h1 { color:#333; }" & vbCrLf & _
                     "    pre { background:#222; color:#eee; padding:15px; border-radius:5px; white-space: pre-wrap; }" & vbCrLf & _
                     "  </style>" & vbCrLf & _
                     "</head>" & vbCrLf & _
                     "<body>" & vbCrLf & _
                     "  <h1>Modulo: " & vbComp.Name & "</h1>" & vbCrLf & _
                     "  <pre>" & Replace(Replace(codice, "<", "&lt;"), ">", "&gt;") & "</pre>" & vbCrLf & _
                     "</body>" & vbCrLf & _
                     "</html>"
        
        ' Salva in file HTML
        htmlPercorso = percorso & vbComp.Name & ".html"
        htmlFile = FreeFile
        Open htmlPercorso For Output As htmlFile
        Print #htmlFile, htmlString
        Close htmlFile
        
        contatoreSalvati = contatoreSalvati + 1 ' Incrementa contatore
        
        ' Logga successo
        Print #logFile, "SUCCESSO - Salvato modulo '" & vbComp.Name & "' in HTML"
        
        On Error GoTo 0
    Next vbComp
    
    ' Scrive la statistica nel log
    Print #logFile, "=== Statistica: moduli salvati con successo = " & contatoreSalvati & " ==="
    
    Print #logFile, "=== Fine esportazione moduli in HTML: " & Now() & " ==="
    Close logFile
    
    MsgBox "Esportazione moduli in HTML completata!" & vbCrLf & _
           "Moduli salvati con successo: " & contatoreSalvati, vbInformation
    Exit Sub

GestioneErrore:
    Print #logFile, "ERRORE - Durante esportazione modulo '" & vbComp.Name & "': " & Err.Description
    Err.Clear
    Resume Next
End Sub




'//======================================================================================================//
'//                         @ESPORTA I MODULI IN FORMATO HTML                      *** FINE ***
'//======================================================================================================//


'//======================================================================================================//
'//                         @ESPORTA LE FORM IN FORMATO HTML                      *** INIZIO ***
'//======================================================================================================//


Sub EsportaTutteLeFormInHTML()
    Dim vbComp As Object ' VBIDE.VBComponent per il modulo di classe della Form
    Dim percorso As String
    Dim logFile As Integer
    Dim htmlFile As Integer
    Dim htmlString As String
    Dim logPercorso As String
    Dim htmlPercorso As String
    Dim contatoreSalvati As Long ' Contatore form salvate
    
    ' Imposta la cartella di esportazione
    percorso = CurrentProject.Path & "\..\OBJECT\FORM\" ' Cartella Form
    logPercorso = percorso & "LOG_EsportaTutteLeFormInHTML.txt"
    
    ' Crea la cartella se non esiste
    If Dir(percorso, vbDirectory) = "" Then MkDir percorso
    
    ' Apre il file di log in modalità append
    logFile = FreeFile
    Open logPercorso For Append As logFile
    Print #logFile, "=== Inizio esportazione Form in HTML: " & Now() & " ==="
    
    contatoreSalvati = 0
    
    ' Scorre tutti i componenti del progetto VBA
    For Each vbComp In Application.VBE.VBProjects(1).VBComponents
        ' Filtra solo i moduli di classe Form (tipo 100 = vbext_ct_ClassModule)
        If vbComp.Type = 100 Then
            ' Considera solo quelli che hanno un nome che corrisponde a Form
            ' Per sicurezza, controlla se esiste la Form con quel nome nel progetto Access
            On Error Resume Next
            Dim frm As Access.Form
            Set frm = Forms(vbComp.Name)
            On Error GoTo 0
            
            ' Se la Form esiste, procedi (oppure usa solo la tipologia modulo di classe)
            ' Nota: non sempre il nome del modulo di classe corrisponde alla Form aperta in Forms() perché la Form potrebbe non essere aperta
            ' Quindi usiamo il nome del componente per esportare il modulo
            
            On Error GoTo GestioneErrore
            
            Dim codice As String
            Dim i As Long
            codice = ""
            For i = 1 To vbComp.CodeModule.CountOfLines
                codice = codice & vbComp.CodeModule.Lines(i, 1) & vbCrLf
            Next i
            
            ' Costruisci contenuto HTML
            htmlString = "<!DOCTYPE html>" & vbCrLf & _
                         "<html lang=""it"">" & vbCrLf & _
                         "<head>" & vbCrLf & _
                         "  <meta charset=""UTF-8"">" & vbCrLf & _
                         "  <title>Form: " & vbComp.Name & "</title>" & vbCrLf & _
                         "  <style>" & vbCrLf & _
                         "    body { font-family: Consolas, monospace; background:#f0f0f0; padding:20px; }" & vbCrLf & _
                         "    h1 { color:#333; }" & vbCrLf & _
                         "    pre { background:#222; color:#eee; padding:15px; border-radius:5px; white-space: pre-wrap; }" & vbCrLf & _
                         "  </style>" & vbCrLf & _
                         "</head>" & vbCrLf & _
                         "<body>" & vbCrLf & _
                         "  <h1>Form: " & vbComp.Name & "</h1>" & vbCrLf & _
                         "  <pre>" & Replace(Replace(codice, "<", "&lt;"), ">", "&gt;") & "</pre>" & vbCrLf & _
                         "</body>" & vbCrLf & _
                         "</html>"
            
            ' Salva in file HTML
            htmlPercorso = percorso & vbComp.Name & ".html"
            htmlFile = FreeFile
            Open htmlPercorso For Output As htmlFile
            Print #htmlFile, htmlString
            Close htmlFile
            
            contatoreSalvati = contatoreSalvati + 1
            
            ' Log successo
            Print #logFile, "SUCCESSO - Salvata Form '" & vbComp.Name & "' in HTML"
            
            On Error GoTo 0
        End If
    Next vbComp
    
    ' Statistica
    Print #logFile, "=== Statistica: Form salvate con successo = " & contatoreSalvati & " ==="
    Print #logFile, "=== Fine esportazione Form in HTML: " & Now() & " ==="
    Close logFile
    
    MsgBox "Esportazione Form in HTML completata!" & vbCrLf & _
           "Form salvate con successo: " & contatoreSalvati, vbInformation
    Exit Sub

GestioneErrore:
    Print #logFile, "ERRORE - Durante esportazione Form '" & vbComp.Name & "': " & Err.Description
    Err.Clear
    Resume Next
End Sub



'//======================================================================================================//
'//                         @ESPORTA LE FORM IN FORMATO HTML                      *** FINE ***
'//======================================================================================================//



'//======================================================================================================//
'//                         @ESPORTA LE MACRO IN FORMATO HTML                      *** INZIO ***
'//======================================================================================================//


Sub EsportaTutteLeMacroInHTML()
    Dim obj As AccessObject
    Dim percorso As String
    Dim logFile As Integer
    Dim htmlFile As Integer
    Dim logPercorso As String
    Dim htmlPercorso As String
    Dim txtPercorso As String
    Dim contatoreSalvati As Long
    Dim testoMacro As String
    Dim fileTxt As Integer
    Dim linea As String  ' <-- dichiarazione variabile
    
    ' Percorso per salvataggio
    percorso = CurrentProject.Path & "\..\OBJECT\MACRO\"
    logPercorso = percorso & "LOG_EsportaTutteLeMacroInHTML.txt"
    
    ' Crea cartella se non esiste
    If Dir(percorso, vbDirectory) = "" Then MkDir percorso
    
    ' Apri file log
    logFile = FreeFile
    Open logPercorso For Append As logFile
    Print #logFile, "=== Inizio esportazione Macro in HTML: " & Now() & " ==="
    
    contatoreSalvati = 0
    
    For Each obj In CurrentProject.AllMacros
        On Error GoTo GestioneErrore
        
        ' Percorso file .txt temporaneo
        txtPercorso = percorso & obj.Name & ".txt"
        
        ' Esporta la macro in formato testo
        Application.SaveAsText acMacro, obj.Name, txtPercorso
        
        ' Leggi il contenuto del file testo
        fileTxt = FreeFile
        Open txtPercorso For Input As fileTxt
        testoMacro = ""
        Do While Not EOF(fileTxt)
            Line Input #fileTxt, linea
            testoMacro = testoMacro & linea & vbCrLf
        Loop
        Close fileTxt
        
        ' Costruisci HTML con il testo della macro
        Dim htmlString As String
        htmlString = "<!DOCTYPE html>" & vbCrLf & _
                     "<html lang=""it"">" & vbCrLf & _
                     "<head>" & vbCrLf & _
                     "  <meta charset=""UTF-8"">" & vbCrLf & _
                     "  <title>Macro: " & obj.Name & "</title>" & vbCrLf & _
                     "  <style>" & vbCrLf & _
                     "    body { font-family: Consolas, monospace; background:#f0f0f0; padding:20px; }" & vbCrLf & _
                     "    h1 { color:#333; }" & vbCrLf & _
                     "    pre { background:#222; color:#eee; padding:15px; border-radius:5px; white-space: pre-wrap; }" & vbCrLf & _
                     "  </style>" & vbCrLf & _
                     "</head>" & vbCrLf & _
                     "<body>" & vbCrLf & _
                     "  <h1>Macro: " & obj.Name & "</h1>" & vbCrLf & _
                     "  <pre>" & Replace(testoMacro, "<", "&lt;") & "</pre>" & vbCrLf & _
                     "</body>" & vbCrLf & _
                     "</html>"
        
        ' Salva il file HTML
        htmlPercorso = percorso & obj.Name & ".html"
        htmlFile = FreeFile
        Open htmlPercorso For Output As htmlFile
        Print #htmlFile, htmlString
        Close htmlFile
        
        ' (Opzionale) elimina il file txt temporaneo
        Kill txtPercorso
        
        contatoreSalvati = contatoreSalvati + 1
        Print #logFile, "SUCCESSO - Salvata macro '" & obj.Name & "' in HTML"
        
        On Error GoTo 0
    Next obj
    
    Print #logFile, "=== Statistica: macro salvate con successo = " & contatoreSalvati & " ==="
    Print #logFile, "=== Fine esportazione Macro in HTML: " & Now() & " ==="
    Close logFile
    
    MsgBox "Esportazione Macro in HTML completata!" & vbCrLf & _
           "Macro salvate con successo: " & contatoreSalvati, vbInformation
    Exit Sub
    
GestioneErrore:
    Print #logFile, "ERRORE - Durante esportazione macro '" & obj.Name & "': " & Err.Description
    Err.Clear
    Resume Next
End Sub



'//======================================================================================================//
'//                         @ESPORTA LE MACRO IN FORMATO HTML                      *** FINE ***
'//======================================================================================================//




'//======================================================================================================//
'//                         @ESPORTA LE TABELLE IN FORMATO HTML                      *** INIZIO ***
'//======================================================================================================//




Sub EsportaStrutturaTutteLeTabelleInHTML()
    ' Dichiarazione delle variabili
    Dim db As DAO.Database             ' Riferimento al database corrente
    Dim rs As DAO.Recordset           ' Recordset per scorrere le tabelle fisiche
    Dim tdf As DAO.TableDef           ' Oggetto per accedere alla struttura della tabella
    Dim fld As DAO.Field              ' Oggetto per accedere ai campi della tabella
    Dim htmlString As String          ' Stringa per contenere il contenuto HTML
    Dim percorso As String            ' Percorso in cui salvare i file HTML
    Dim htmlFile As Integer           ' Identificatore file per scrittura
    Dim htmlPercorso As String        ' Percorso completo del file HTML da scrivere
    Dim nomeTabella As String         ' Nome della tabella corrente
    Dim countSalvati As Long          ' Contatore delle tabelle esportate
    Dim sql As String                 ' Query SQL per ottenere tabelle fisiche
    
    ' Imposta il percorso in cui verranno salvati i file HTML
    percorso = CurrentProject.Path & "\..\OBJECT\TABLES\"
    
    ' Verifica se la cartella esiste, altrimenti la crea
    If Dir(percorso, vbDirectory) = "" Then
        MkDir percorso
        Debug.Print "Cartella creata: " & percorso
    Else
        Debug.Print "Cartella già esistente: " & percorso
    End If

    Set db = CurrentDb()              ' Apre riferimento al database corrente
    countSalvati = 0                  ' Inizializza contatore esportazioni
    
    ' Query sulle tabelle di sistema per estrarre SOLO le tabelle fisiche utente
    sql = "SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' Cicla tutte le tabelle fisiche trovate
    Do While Not rs.EOF
        nomeTabella = rs!Name         ' Recupera nome della tabella corrente
        On Error GoTo ErrHandler      ' Gestione errori

        ' Ottiene definizione tabella
        Set tdf = db.TableDefs(nomeTabella)
        
        ' Costruzione intestazione HTML
        htmlString = "<!DOCTYPE html>" & vbCrLf & _
                     "<html lang=""it"">" & vbCrLf & _
                     "<head>" & vbCrLf & _
                     "  <meta charset=""UTF-8"">" & vbCrLf & _
                     "  <title>Struttura Tabella: " & nomeTabella & "</title>" & vbCrLf & _
                     "  <style>" & vbCrLf & _
                     "    body { font-family: Consolas, monospace; background:#f0f0f0; padding:20px; }" & vbCrLf & _
                     "    h1 { color:#333; }" & vbCrLf & _
                     "    table { border-collapse: collapse; width: 100%; max-width: 600px; }" & vbCrLf & _
                     "    th, td { border: 1px solid #666; padding: 8px; text-align: left; }" & vbCrLf & _
                     "    th { background-color: #ccc; }" & vbCrLf & _
                     "  </style>" & vbCrLf & _
                     "</head>" & vbCrLf & _
                     "<body>" & vbCrLf & _
                     "  <h1>Struttura Tabella: " & nomeTabella & "</h1>" & vbCrLf & _
                     "  <table>" & vbCrLf & _
                     "    <tr><th>Nome Campo</th><th>Tipo</th></tr>" & vbCrLf
        
        ' Aggiunge righe HTML per ogni campo nella tabella
        For Each fld In tdf.Fields
            htmlString = htmlString & "    <tr><td>" & fld.Name & "</td><td>" & TipoCampoHTML(fld) & "</td></tr>" & vbCrLf
        Next fld
        
        ' Chiude il corpo HTML
        htmlString = htmlString & "  </table>" & vbCrLf & _
                     "</body>" & vbCrLf & _
                     "</html>"
        
        ' Salvataggio del file HTML
        htmlPercorso = percorso & nomeTabella & ".html"
        htmlFile = FreeFile                          ' Ottiene numero file disponibile
        Open htmlPercorso For Output As htmlFile     ' Apre file in scrittura
        Print #htmlFile, htmlString                  ' Scrive contenuto HTML
        Close htmlFile                               ' Chiude file
        
        countSalvati = countSalvati + 1              ' Incrementa contatore
        
        rs.MoveNext                                  ' Passa alla prossima tabella
        GoTo ContinueLoop
        
              
' --- GESTIONE ERRORI PER TABELLASINGOLA ---
ErrHandler:
        Debug.Print "Errore con tabella: " & nomeTabella & " - " & Err.Description
        Resume Next

ContinueLoop:
    Loop
    
    ' Pulizia oggetti DAO
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' Messaggio finale all'utente
    MsgBox "Esportazione completata!" & vbCrLf & _
           "Tabelle esportate: " & countSalvati & vbCrLf & _
           "Cartella: " & percorso, vbInformation
End Sub


Public Function TipoCampoHTML(fld As DAO.Field) As String
    ' Ritorna una stringa testuale corrispondente al tipo del campo DAO
    Select Case fld.Type
        Case dbBoolean: TipoCampoHTML = "YESNO"
        Case dbByte: TipoCampoHTML = "BYTE"
        Case dbInteger: TipoCampoHTML = "SMALLINT"
        Case dbLong: TipoCampoHTML = "LONG"
        Case dbCurrency: TipoCampoHTML = "CURRENCY"
        Case dbSingle: TipoCampoHTML = "SINGLE"
        Case dbDouble: TipoCampoHTML = "DOUBLE"
        Case dbDate: TipoCampoHTML = "DATETIME"
        Case dbText: TipoCampoHTML = "TEXT(" & fld.Size & ")"
        Case dbLongBinary: TipoCampoHTML = "OLEOBJECT"
        Case dbMemo: TipoCampoHTML = "MEMO"
        Case Else: TipoCampoHTML = "UNKNOWN"
    End Select
End Function






'//======================================================================================================//
'//                         @ESPORTA LE TABELLE IN FORMATO HTML                      *** FINE ***
'//======================================================================================================//

'// CREAZIONE DI UNA QUERY CREATE CHE RIGENERA LE TABELLE SALVATE
'//---------------------------------------------------------------------------------------------//


Public Function TipoCampoDDL(fld As DAO.Field) As String
    Select Case fld.Type
        Case dbBoolean
            TipoCampoDDL = "YESNO"
        Case dbByte
            TipoCampoDDL = "BYTE"
        Case dbInteger
            TipoCampoDDL = "INTEGER"
        Case dbLong
            TipoCampoDDL = "LONG"
        Case dbCurrency
            TipoCampoDDL = "CURRENCY"
        Case dbSingle
            TipoCampoDDL = "SINGLE"
        Case dbDouble
            TipoCampoDDL = "DOUBLE"
        Case dbDate
            TipoCampoDDL = "DATETIME"
        Case dbText
            TipoCampoDDL = "TEXT(" & fld.Size & ")"
        Case dbLongBinary
            TipoCampoDDL = "OLEOBJECT"
        Case dbMemo
            TipoCampoDDL = "MEMO"
        Case Else
            TipoCampoDDL = "UNKNOWN"
    End Select
End Function

'//ESPORTA TABELLA SOTTOFORM DI CREATE QUERY  @CREATE@TABLE_(genero una query CREATE di @rigenerazione della tabella)
'//Note: questa funzione crea una query di CREAZIONE TABELLA in formato html.
Sub EsportaTABELLAdaCreareConLaQueryCreateHTML()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim htmlString As String
    Dim sqlCreate As String
    Dim htmlFile As Integer
    Dim percorso As String
    Dim percorsoFile As String
    Dim nomeTabella As String
    Dim prima As Boolean
    Dim count As Long
    
    percorso = CurrentProject.Path & "\..\OBJECT\TABLES\"
    If Dir(percorso, vbDirectory) = "" Then MkDir percorso
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0", dbOpenSnapshot)
    
    count = 0
    Do While Not rs.EOF
        nomeTabella = rs!Name
        Set tdf = db.TableDefs(nomeTabella)
        
        sqlCreate = "CREATE TABLE " & nomeTabella & " ("
        prima = True
        
        For Each fld In tdf.Fields
            If Not prima Then sqlCreate = sqlCreate & ", "
            sqlCreate = sqlCreate & fld.Name & " " & TipoCampoDDL(fld)
            prima = False
        Next fld
        
        sqlCreate = sqlCreate & ");"
        
        htmlString = "<!DOCTYPE html>" & vbCrLf & _
                     "<html lang=""it"">" & vbCrLf & _
                     "<head><meta charset=""UTF-8""><title>Query CREATE TABLE: " & nomeTabella & "</title>" & _
                     "<style>body {font-family: Consolas; background:#f0f0f0; padding:20px;} pre {background:#fff; border:1px solid #ccc; padding:10px;}</style>" & _
                     "</head><body>" & vbCrLf & _
                     "<h1>Query CREATE TABLE: " & nomeTabella & "</h1>" & vbCrLf & _
                     "<pre>" & sqlCreate & "</pre>" & vbCrLf & _
                     "</body></html>"
        
        percorsoFile = percorso & nomeTabella & "_TabellaDaRigenerareConLaQUERY_CREATE.html"
        htmlFile = FreeFile
        Open percorsoFile For Output As htmlFile
        Print #htmlFile, htmlString
        Close htmlFile
        
        Debug.Print "DDL salvata: " & nomeTabella
        count = count + 1
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    MsgBox "DDL CREATE esportato in HTML per " & count & " tabelle.", vbInformation
End Sub


