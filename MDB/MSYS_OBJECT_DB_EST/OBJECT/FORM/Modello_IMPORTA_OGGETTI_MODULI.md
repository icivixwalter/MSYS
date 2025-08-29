# Modello_IMPORTA_OGGETTI_MODULI.md


'//==========================================================================================================//
'// IMPORTA GLI OGGETTI MODULI *** INIZIO *** '@IMPORT@MODULI
'//==========================================================================================================//


Private Sub Cmd_Importa_MODULI_Click()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim modulesCollection As Collection
    Dim modulesCollection_v As Variant
    Dim moduleExists As Boolean
    Dim deletedModulesCount_i As Integer
    Dim deletetModulesCount_i As Integer
    Dim deletedModules_s As String
    Dim importedCount As Integer
    Dim importedModules_s As String
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo Err_Cmd_Import_MODULI_Click
   
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
    
    ' Recupero la path e il file dal database esterno
    sourceDBPath = PathFile_s_pFunct
    
    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE: path e file sono null, uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        Exit Sub
    End If
    
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO LE FORM DA QUESTO DATABASE; path e del file: " & sourceDBPath, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Ciclo nella collezione per il controllo dei moduli precaricati
    For Each modulesCollection_v In modulesCollection
        Debug.Print "I MODULI nella collezione: " & modulesCollection_v
        
        ' Verifica se il modulo esiste nel database corrente
        moduleExists = False
        Set dbs = Application.CurrentProject
        For Each obj In dbs.AllModules
            If obj.Name = modulesCollection_v Then
                moduleExists = True
                Exit For
            End If
        Next obj
        
        If moduleExists Then
            ' Se il modulo esiste nel database corrente, cancellalo
            DoCmd.DeleteObject acModule, modulesCollection_v
            deletetModulesCount_i = deletetModulesCount_i + 1
            deletedModules_s = deletedModules_s & modulesCollection_v & vbCrLf
            Debug.Print "Modulo eliminato: " & modulesCollection_v
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
            
                '//APRO IL DB ESTERNO su eseguot una query sql sulla tabella di sistema estraendo un recordset superando due filtri _
                    il PRIMO FILTRO     : -32761 il codice dei moduli (estraggo solo i moduli) _
                    il SECONDO FILETRO  : il nome del modulo recuperato dalla collection (se esisten nella tabella di sistema)
                moduleExists = False
                Set rs = externalDB.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32761 AND Name='" & modulesCollection_v & "'")
                '//SE IL RS E' POPOLATO vuol dire che nella tabella di sistema, del progetto mdb esiste il modulo ricercato
                If Not rs.EOF Then
                    '//imposto a true = trovato modulo
                    moduleExists = True
                End If
                rs.Close
                
                '//TRUE = esiste il modulo e lo importo
                If moduleExists Then
                    ' Se il modulo esiste nel database esterno, importalo
                    DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acModule, modulesCollection_v, modulesCollection_v
                    importedCount = importedCount + 1
                    importedModules_s = importedModules_s & modulesCollection_v & vbCrLf
                    Debug.Print "Modulo importato: " & modulesCollection_v
                Else
                    ' Modulo non trovato nel database esterno messaggio print.
                    Debug.Print "Modulo non trovato nel database esterno: " & modulesCollection_v
                End If
            '//----------------------------------------------------------------------------------------------------------------------//
        
    Next modulesCollection_v
    
    ' Chiudo il database esterno
    externalDB.Close
    
    ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO FORM PRECARICATE DA IMPORTARE:" & vbCrLf & _
               "Modules importate: " & importedCount & vbCrLf & importedModules_s & _
               "Modules esistenti cancellate: " & deletetModulesCount_i & vbCrLf & deletedModules_s, vbInformation
    End If
    
                  '//ciclo nella collezione per il controllo delle MODULI precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     For Each obj In dbs.AllModules
                            '//vado all'ultima query
                        DoCmd.SelectObject acModule, obj.Name, True
                    
                     Next obj
                '//--------------------------------------------------------------------------------------//

    
    

Exit_Cmd_Import_MODULI_Click:
    Exit Sub

Err_Cmd_Import_MODULI_Click:
    MsgBox Err.Description & " - Error Message -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Import_MODULI_Click

End Sub



'//==========================================================================================================//
'// IMPORTA GLI OGGETTI MODULI *** INIZIO *** '@IMPORT@MODULI
'//==========================================================================================================//


