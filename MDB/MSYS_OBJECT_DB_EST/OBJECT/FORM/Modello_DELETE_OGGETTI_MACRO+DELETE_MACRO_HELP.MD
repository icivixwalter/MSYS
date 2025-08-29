# Modello_DELETE_OGGETTI_MACRO+DELETE_MACRO_HELP


'//=====================================================================================================================//
'//                     DELETE OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE MACRO HELP
'//=====================================================================================================================//
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell'importazione.
'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se le MACRO esistono _
        nel database corrente mediante il confronto tra : _
        LA COLLECTION MACRO = GLI OGGETTI MACRO che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@macro@caricati_(cancellazione delle MACRO già caricate nel db esterno) _
        @cancella@MACRO@esterne_(attenzione puoi cancellare le MACRO nel @db@esterno)

'//FAQ : _
        DOVE SI CANCELLANO LE MACRO PRECARICATE _
        COME CANCELLARE LE MACRO ESTERNE _
        LE MACRO ESTERNE DA CANCELLARE
 
'// PARAMETRI DA CAMBIARE : _
    -----> MACRO _
    -----> DELLE@MACRO _
    -----> DELETE_MACRO _
    -----> CollectionMacro_PFunct _
    -----> macrosCollection _
    -----> objectMacro _
    -----> MATRICE_MACRO_Names_s _
    -----> macrosCount_i _
    -----> MacroDef _
    -----> mdf _


Private Sub Cmd_Delete_MACRO_Click()

    Dim dbCurrent As DAO.Database
    Dim dbs As Object

    Dim obj As AccessObject
    
    Dim macrosCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_MACRO_Names_s() As String             'matrice di MACRO
    Dim tempCount As Integer
    Dim macrosCount_i As Integer
    Dim od As AccessObject
    Dim icount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_MACRO_Click

    '//RESET
    icount = 0
    deletetMacrosCount_i = 0
    deletedMacros_s = ""

    
    
    
    
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
                
    
    
    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION MACRO per ottenere la collezione di MACRO

    Set macrosCollection = CollectionMacro_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutte le MACRO nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE MACRO
    icount = CurrentProject.AllMacros.Count
    

    '//se ci sono MACRO da cancellare
    If icount > 0 Then

        '//reimposto la matrice sulle MACRO esistenti nel db corrente
        ReDim MATRICE_MACRO_Names_s(icount - 1)

        ' Costruisci un array con i nomi delle MACRO
        tempCount = 0

            '//CARICO LA MATRICE: ciclo for per caricare MATRICE_MACRO_Names_s(tempCount)
            Set dbs = Application.CurrentProject
            ' Ricerca oggetti AccessObject aperti in insieme AllMacros.
            For Each obj In dbs.AllMacros
                    ' Stampa nome di obj.
                    Debug.Print obj.Name
                    MATRICE_MACRO_Names_s(tempCount) = obj.Name
                    tempCount = tempCount + 1
            Next obj
    
    
        

        ' Cancella le MACRO usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione MACRO: " & MATRICE_MACRO_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionMacrosDELETE_Funct_b(macrosCollection, MATRICE_MACRO_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello la MACRO trovata sia nella collezione che nella matrice : " & MATRICE_MACRO_Names_s(i)
                DoCmd.DeleteObject acMacro, MATRICE_MACRO_Names_s(i)

                If Err.Number <> 0 Then
                    Debug.Print "Errore durante la cancellazione della MACRO: " & MATRICE_MACRO_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetMacrosCount_i = deletetMacrosCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "Le MACRO CANCELLATE DA QUESTO DATABASE SONO :  " & deletetMacrosCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_MACRO_Click:
    Exit Sub

Err_Cmd_Delete_MACRO_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_MACRO_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione di MACRO
Function InCollectionMacrosDELETE_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionMacrosDELETE_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionMacrosDELETE_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionMacrosDELETE_Funct_b = False
    On Error GoTo 0
End Function



Private Sub Cmd_Delete_MACRO_HELP_Click()

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
            MyFile_s = "HELP_Cmd_Delete_MACRO.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//


End Sub


'//=====================================================================================================================//
'//                     DELETE OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE MACRO HELP
'//=====================================================================================================================//

