'//    MODULO ---> NAME ESPORTA_Mdl01_TUTTI_GLI_OGGETTI_IN_FILE_LOG.MD
'//****************************************************************************//'
'// ATTIVITA : esporto i seguenti oggetti in un file .txt
'//      01) @esporta@query_(la procedura ora esporta le query in un file.txt)
'//          @esportare@query_(in un file .log ed in html)
'//   si trova nel file : MSys_MODULI.mdb
'//
'//   apri db: START "APRI MODULI" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\MSys_MODULI.mdb"
'//   apri db: START "APRI PATH" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\"
'//
'//   faq:
'//      come esporto le query; con quale procedura esporto le query
'//      come esportare le query del db corrente; modulo per l'esportazione delle query
'//
'//*******************************************************************************************************//



Option Compare Database
Option Explicit






'//======================================================================================================//
'//                         @ESPORTA_NOME_QUERY_IN_LOG_EsportaTutteLeQueryInHTML.txt (Esporta in HTML)                      *** INIZIO ***
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
    
    ' Recupera la cartella corrente del database
    percorso = CurrentProject.PATH & "\"
    logPercorso = percorso & "LOG_EsportaTutteLeQueryInHTML.txt"
    
    ' Apre il file di log in modalità append per tracciare l'esportazione
    logFile = FreeFile
    Open logPercorso For Append As logFile
    Print #logFile, "=== Inizio esportazione SQL in HTML: " & Now() & " ==="
    
    ' Apre il database corrente
    Set db = CurrentDb
    
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
                         "  <pre>" & Replace(qdf.SQL, "<", "&lt;") & "</pre>" & vbCrLf & _
                         "</body>" & vbCrLf & _
                         "</html>"
            
            ' Salva il file HTML con nome della query nella cartella del DB
            htmlPercorso = percorso & qdf.Name & ".html"
            htmlFile = FreeFile
            Open htmlPercorso For Output As htmlFile
            Print #htmlFile, htmlString
            Close htmlFile
            
            ' Scrive nel log il successo dell'esportazione
            Print #logFile, "SUCCESSO - Salvata query '" & qdf.Name & "' in HTML"
            
            On Error GoTo 0  ' reset error handling per il prossimo ciclo
        End If
    Next qdf
    
    ' Segnala la fine del processo di esportazione nel log
    Print #logFile, "=== Fine esportazione SQL in HTML: " & Now() & " ==="
    Close logFile
    
    ' Pulisce variabili per evitare perdite di memoria
    Set qdf = Nothing
    Set db = Nothing
    
    MsgBox "Esportazione SQL in HTML completata!", vbInformation
    Exit Sub
    
GestioneErrore:
    Print #logFile, "ERRORE - Durante esportazione query '" & qdf.Name & "': " & Err.Description
    Err.Clear
    Resume Next
End Sub
'//======================================================================================================//
'//                         @ESPORTA_NOME_QUERY_IN_LOG_EsportaTutteLeQueryInHTML.txt (Esporta in HTML)                      *** FINE ***
'//======================================================================================================//





