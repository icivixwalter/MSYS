# MODELLO_DELETE_REPORT_OLD.MD



'//DELETE REPORT CON INDICE NUMERICO
'//==========================================================================================================//
'// CANCELLA GLI OGGETTI REPORT *** INIZIO *** '@DELETE@REPORT
'// ATTENZIONE = Al posto di usare AllReports il quale crea problemi di cancellazione perchè _
        sembra andare fuori indice si è utilizzato la proprieta dbs.AllReports.Count in modo da _
        recuperare il numero dei report nel db corrente ed attivare una iterazione basata sull'indice _
        totale dei report salvati. In questo modo non si hanno errore di cancellazione.
'//==========================================================================================================//
Private Sub Cmd_Delete_REPORT_Click()
    Dim Rptobj As AccessObject
    Dim dbs As Object
    Dim dbCurrent As DAO.Database
    Dim tdf As DAO.TableDef
    
    Dim objectTables As Collection
    Dim varItem As Variant
    Dim reportNames() As String             'matrice di report
    Dim i As Integer, reportCount As Integer
    Dim tempCount As Integer
    
'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_REPORT_Click
    
    '//RESET
    deletetReportesCount_i = 0
    myreportName_s = ""
        
    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION REPORT per ottenere la collezione di REPORT
    Set reportsCollection = CollectionReport_PFunct()
    '//------------------------------------------------------------------------//
    
    ' Loop attraverso tutte I REPORT nel database corrente
    Set dbs = Application.CurrentProject
    
    '//CALCOLO IL NUMERO DI INDICE DEI REPORT
    reportCount = dbs.AllReports.Count
    
    '//SE INDICE REPORT > 0 ESEGUO LA MATRICE DEI REPORT
    If reportCount > 0 Then
    
        '//REIMPOSTO LA MATRICE REPORT
        ReDim reportNames(reportCount - 1)
        
        ' Costruisci un array con i nomi dei report
        tempCount = 0
        
        '//CARICO LA MATRICE DEI REPORT CORRENTI reportNames(tempCount)
        For Each Rptobj In dbs.AllReports
            reportNames(tempCount) = Rptobj.Name
            tempCount = tempCount + 1
        Next Rptobj
        
        '//CICLO NELLA MATRICE DEI REPORT
        ' Cancella i report usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione report: " & reportNames(i)
            
            DoCmd.DeleteObject acReport, reportNames(i)
            If Err.Number <> 0 Then
                Debug.Print "Errore durante la cancellazione del report: " & reportNames(i) & " - " & Err.Description
                Err.Clear  ' Pulisci l'errore per la prossima iterazione
            Else
                deletetReportesCount_i = deletetReportesCount_i + 1
            End If
            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If
    
    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "I REPORT CANCELLATI DA QUESTO DATABASE SONO :  " & deletetReportesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_REPORT_Click:
    Exit Sub

Err_Cmd_Delete_REPORT_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_REPORT_Click

End Sub

'//==========================================================================================================//
'// CANCELLA GLI OGGETTI REPORT *** INIZIO *** '@DELETE@REPORT  *** FINE ****
'//==========================================================================================================//

