# Modello_DELETE_OGGETTI_MODULI.md


'//=====================================================================================================================//
'//                     DELETE OGGETTI MODULI DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE MODULI HELP
'//=====================================================================================================================//

Private Sub Cmd_Delete_MODULI_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim modulesCollection As Collection
    'Dim modulesCollection_v As Variant
    Dim moduleExists As Boolean
    'Dim deletedModulesCount_i As Integer
    'Dim deletetModulesCount_i As Integer
    'Dim deletedModules_s As String
    'Dim importedModules_s As String
    
    On Error GoTo Err_Cmd_Delete_MODULI_Click
   
    ' Reset count
    deletedModulesCount_i = 0
    deletetModulesCount_i = 0
    deletedModules_s = ""
    importedCount = 0
    importedModules_s = ""
     
     
     
     
     

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
         
     
    ' Recupero la collezione moduli da importare
    Set modulesCollection = CollectionModuli_PFunct()
                        
    ' Ciclo nella collezione per il controllo delle form precaricate
    For Each modulesCollection_v In modulesCollection
        Debug.Print "I MODULI nella collezione: " & modulesCollection_v
        
        ' Verifica se il modulo esiste nel database
        moduleExists = False
        Set dbs = Application.CurrentProject
        
        '//itero nell'insime ALLMODULES
        For Each obj In dbs.AllModules
        
            '//CONFRONTA SE ESISTE IMPOSTA TRUE
            If obj.Name = modulesCollection_v Then
                moduleExists = True
                Exit For
            End If
        Next obj
        
        '//TRUE= ESISTE E CANCELLO IL MODULO
        If moduleExists Then
            ' Se il modulo esiste, cancellalo
            DoCmd.DeleteObject acModule, modulesCollection_v
            deletetModulesCount_i = deletetModulesCount_i + 1
            deletedModules_s = deletedModules_s & modulesCollection_v & vbCrLf
            Debug.Print "Modulo eliminato: " & modulesCollection_v
        Else
            ' Se il modulo non esiste, considera per l'importazione
            importedModulesCount_i = importedModulesCount_i + 1
            importedModules_s = importedModules_s & modulesCollection_v & vbCrLf
        End If
    Next modulesCollection_v
                        
    ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO MODULI PRECARICATE DA IMPORTARE:" & vbCrLf & _
               "Modules importate: " & importedModulesCount_i & vbCrLf & importedModules_s & _
               "Modules esistenti cancellate: " & deletetModulesCount_i & vbCrLf & deletedModules_s, vbInformation
    End If
    
     ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
   
        ' Optional: display a message indicating that the modules have been deleted
        MsgBox "Tutti i moduli standard sono stati cancellati. Totale moduli cancellati: " & deletetModulesCount_i, vbInformation, "Moduli Cancellati"
    
    End If

Exit_Cmd_Delete_MODULI_Click:
    Exit Sub

Err_Cmd_Delete_MODULI_Click:
    MsgBox Err.Description & " - Error Message -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Delete_MODULI_Click

End Sub

'//DELETE MODULI        ***FINE ***
'//==========================================================================================================//



'//Cmd_DELETE_MODULI_HELP_Click  *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_DELETE_MODULI_HELP_Click()

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
            MyFile_s = "HELP_Cmd_Delete_MODULI.pdf "
               
               
             '//CHIAMO LA FUNZIONE INTERNA PER L'APERTURA DEI FILE
                Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
                 
           
    '//--------------------------------------------------------------------------------//--------//




End Sub

'//Cmd_DELETE_MODULI_HELP_Click  *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=====================================================================================================================//
'//                     DELETE OGGETTI MODULI DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE MODULI HELP
'//=====================================================================================================================//






