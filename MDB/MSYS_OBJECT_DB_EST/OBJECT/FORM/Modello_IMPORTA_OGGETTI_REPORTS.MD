# Modello_IMPORTA_OGGETTI_REPORTS.MD




Private Sub Cmd_Importa_REPORT_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim reportsCollection As Collection
    Dim reportExists As Boolean
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo Err_Cmd_Import_REPORT_Click
   
    ' Reset counts and strings
    deletetReportesCount_i = 0
    importedReportesCount_i = 0
    importedReports_s = ""
    
    
    
    
    
            '//=====================================================================================================================//
                '//                             STEP_02 = CONTROLLO PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYSDbEst_GE_OBJECT.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = CONTROLLO PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

    
    
    
    
    
    
    
    
    
    
    
    
    
    ' Recupero la collezione di report da importare
    Set reportsCollection = CollectionReport_PFunct()
    
    ' Recupero il percorso e il nome del file del database esterno
    sourceDBPath = PathFile_s_pFunct
    
    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE: il percorso del file del database esterno non è stato specificato. Uscita dalla routine.", vbCritical, "Errore: Percorso del File Non Specificato"
        Exit Sub
    End If
    
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO I REPORT DA QUESTO DATABASE; Percorso del file: " & sourceDBPath, vbExclamation, "Informazione: Chiamata Funzione PathFile"
    End If
    
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Ciclo nella collezione per il controllo dei report precaricati
    For Each reportName In reportsCollection
        Debug.Print "Report nella collezione: " & myreportName_s
        
        ' Verifica se il report esiste nel database corrente
        reportExists = False
        Set dbs = Application.CurrentProject
        For Each obj In dbs.AllReports
            If obj.Name = reportName Then
                reportExists = True
                Exit For
            End If
        Next obj
        
        If reportExists Then
            ' Se il report esiste nel database corrente, cancellalo
            DoCmd.DeleteObject acReport, reportName
            deletetReportesCount_i = deletetReportesCount_i + 1
            Debug.Print "Report eliminato: " & reportName
        End If
        
        '//CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects
        '//----------------------------------------------------------------------------------------------------------------------//
        ' Verifica se il modulo esiste nel database esterno
        ' La riga SELECT Name FROM MSysObjects WHERE Type=-32761 è una query SQL utilizzata per selezionare i
        ' nomi degli oggetti di tipo modulo nel database di Access. Ecco una spiegazione dettagliata:

        ' MSysObjects: è una tabella di sistema in Microsoft Access che contiene informazioni su tutti gli oggetti nel database, _
          come tabelle, query, moduli, report, ecc. Type=-32761: questo filtro nella clausola WHERE specifica il tipo di oggetto _
          da selezionare. In Access, ogni tipo di oggetto ha un valore numerico associato nella tabella MSysObjects. _
          Il valore -32761 è il codice che rappresenta i moduli (VBA Modules).
         'Quindi, la query SELECT Name FROM MSysObjects WHERE Type=-32761 seleziona tutti i nomi degli oggetti nella tabella _
         di sistema MSysObjects che sono identificati come moduli VBA.
         ' Verifica se il report esiste nel database esterno
        ' Utilizzo una query SQL per verificare l'esistenza del report nel database esterno
        ' il codice del report è : -32764 controllato con la funzione codici esterni.
        

        reportExists = False
        Set rs = externalDB.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type= -32764 AND Name='" & reportName & "'")
            If Not rs.EOF Then
                reportExists = True
            End If
                rs.Close
            
            If reportExists Then
                ' Se il report esiste nel database esterno, importalo
                DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acReport, reportName, reportName
                importedReportesCount_i = importedReportesCount_i + 1
                importedReports_s = importedReports_s & reportName & vbCrLf
                Debug.Print "Report importato: " & reportName
            Else
                ' Report non trovato nel database esterno
                Debug.Print "Report non trovato nel database esterno: " & reportName
            End If
        Next reportName
        
            ' Chiudo il database esterno
            externalDB.Close
        '//----------------------------------------------------------------------------------------------------------------------//
        
    ' Messaggio di riepilogo finale
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO REPORT PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Report importati: " & importedReportesCount_i & vbCrLf & importedReports_s & _
               "Report esistenti cancellati: " & deletetReportesCount_i, vbInformation
    End If
    
    
                '//ciclo nella collezione per il controllo delle QUERY precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     For Each obj In dbs.AllReports
                            '//vado all'ultima query
                        DoCmd.SelectObject acReport, obj.Name, True
                    
                     Next obj
                '//--------------------------------------------------------------------------------------//


Exit_Cmd_Import_REPORT_Click:
    Exit Sub

Err_Cmd_Import_REPORT_Click:
    MsgBox Err.Description & " - Errore: " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Import_REPORT_Click

End Sub


'//FUNZIONE DI CONTROLLO DEGLI OGGETTI REPORT DEL DB MSYSDBEST *** INIZIO ***
'//----------------------------------------------------------------------------------------//

'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO = todo: modificata
Function ObjectReportExists(db As DAO.Database, objType As Integer, objName As String) As Boolean
    'Dim obj As DAO.Document
    'On Error Resume Next
    'Set obj = db.Containers("Reports").Documents(objName)
    'ObjectReportExists = (Err.Number = 0)
    'Set obj = Nothing
    'On Error GoTo 0
    
    Dim obj As AccessObject
    On Error Resume Next
    For Each obj In CurrentProject.AllReports
        Debug.Print "report corrente controllato :" & obj.Name
        If obj.Name = objName Then
            ReportExistsInCurrentDb = True
            Exit Function
        End If
    Next obj
    ReportExistsInCurrentDb = False
    On Error GoTo 0
    
    
    
End Function


'//FUNZIONE DI CONTROLLO DEGLI OGGETTI REPORT DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//

'//==========================================================================================================//
'// IMPORTA GLI OGGETTI REPORT *** FINE *** '@IMPORT@REPORT
'//==========================================================================================================//

