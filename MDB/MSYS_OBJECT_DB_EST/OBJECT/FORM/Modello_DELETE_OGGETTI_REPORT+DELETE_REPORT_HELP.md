# Modello_DELETE_OGGETTI_REPORT+DELETE_REPORT_HELP.MD

'//=====================================================================================================================//
'//                     DELETE OGGETTI REPROT DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE REPORT HELP
'//=====================================================================================================================//



'//CANCELLA GLI OGGETTI REPORT *** INIZIO *** DELETE_REPORT
'//======================================================================================================//
'//@GESTIONE@DEI@REPORT    (@MODELLO@DELETE@OGGETTI)

'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se i REPORT esistono _
        nel database corrente mediante il confronto tra : _
        LA COLLECTION REPORT = GLI OGGETTI REPORT che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@report@caricati_(cancellazione dei REPORT già caricati nel db esterno) _
        @cancella@TABLE@esterne_(attenzione puoi cancellare i REPORT nel @db@esterno)

'//FAQ : _
        DOVE SI CANCELLANO I REPORT PRECARICATI _
        COME CANCELLARE I REPORT ESTERNI _
        I REPORT ESTERNI DA CANCELLARE

'// PARAMETRI DA CAMBIARE : _
    -----> REPORT _
    -----> DEI@REPORT _
    -----> DELETE_REPORT _
    -----> CollectionReport_PFunct _
    -----> reportsCollection _
    -----> objectReport _
    -----> MATRICE_REPORT_Names_s _
    -----> reportsCount_i _
    -----> ReportDef _
    -----> rpt _

Private Sub Cmd_Delete_REPORT_Click()

    Dim dbCurrent As DAO.Database
    Dim rpt As Report
    Dim reportsCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_REPORT_Names_s() As String             'matrice di REPORT
    Dim tempCount As Integer
    Dim reportsCount_i As Integer
    Dim rdf As AccessObject
    Dim icount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_REPORT_Click

    '//RESET
    icount = 0
    deletetReportsCount_i = 0
    deletedReports_s = ""
    
    
    
    
    

    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION REPORT per ottenere la collezione di REPORT

    Set reportsCollection = CollectionReport_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutti i REPORT nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE REPORT
    icount = Application.CurrentProject.AllReports.Count

    '//se ci sono REPORT da cancellare
    If icount > 0 Then

        '//reimposto la matrice sui REPORT esistenti nel db corrente
        ReDim MATRICE_REPORT_Names_s(icount - 1)

        ' Costruisci un array con i nomi dei REPORT
        tempCount = 0

        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_REPORT_Names_s(tempCount)
        For i = 0 To icount - 1
            Set rdf = Application.CurrentProject.AllReports(i)
            MATRICE_REPORT_Names_s(tempCount) = rdf.Name
            tempCount = tempCount + 1
        Next i

        ' Cancella i REPORT usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione REPORT: " & MATRICE_REPORT_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionREPORTes_Funct_b(reportsCollection, MATRICE_REPORT_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello il REPORT trovato sia nella collezione che nella matrice : " & MATRICE_REPORT_Names_s(i)
                DoCmd.DeleteObject acReport, MATRICE_REPORT_Names_s(i)

                If Err.Number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del REPORT: " & MATRICE_REPORT_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetReportsCount_i = deletetReportsCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "I REPORT CANCELLATI DA QUESTO DATABASE SONO :  " & deletetReportsCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA E GESTIONE ERRORI
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

' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionREPORTes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionREPORTes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionREPORTes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionREPORTes_Funct_b = False
    On Error GoTo 0
End Function


'//CANCELLA GLI OGGETTI REPORT *** FINE *** DELETE_REPORT
'//=========================================================================================================

'//DELETE REPORT HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Delete_REPORT_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_Cmd_DELETE_REPORT.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//DELETE REPORT HELP *** FINE ***
'//-----------------------------------------------------------------------------------//




'//=====================================================================================================================//
'//                     DELETE OGGETTI REPROT DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE REPORT HELP
'//=====================================================================================================================//


