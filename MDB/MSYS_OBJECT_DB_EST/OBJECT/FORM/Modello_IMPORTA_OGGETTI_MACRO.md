


'//=====================================================================================================================//
'//                            IMPORTO LE MACRO *** INIZIO ***
'//
'//=====================================================================================================================//



'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell'importazione.


'//PARAMETRI_DA_CAMBIARE _
        ----> Cmd_Importa_MACRO_Click _
        ----> Dim qdf As DAO.MacroDef _
        ----> Dim strSourceDbPath As String _
        ----> Dim objectMacros As Collection _
        ----> Dim objectExistsInCollectionMacros_Funct_b As Boolean _
        ----> MACRO _
        ----> MACRO _
        ----> COLLECTION MACRO _
        ----> importedMACROesCount_i _
        ----> deleteMACROCount_i _
        ----> deleteNameMACRO_s _


Private Sub Cmd_Importa_MACRO_Click()

    Dim dbCurrent As DAO.Database
    Dim dbSource As DAO.Database
    Dim dbs As Object
    Dim obj As AccessObject

    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strSourceDbPath As String
    Dim objectMacros As Collection
    Dim varItem As Variant
    Dim objectExistsInCollectionMacros_Funct_b As Boolean
    Dim importedMACROesCount_i As Integer
    Dim deleteMACROCount_i As Integer
    Dim deleteNameMACRO_s As String
    Dim importedNameMACRO_s As String
    Dim myNameMACRO_s As String
    Dim Bool1 As Boolean
    Dim MyValue As Integer

    On Error GoTo ErrorHandler

    ' Inizializza le variabili
    importedMACROesCount_i = 0
    deleteMACROCount_i = 0
    deleteNameMACRO_s = ""
    importedNameMACRO_s = ""
    
    
    
    
    
    
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

    

    ' ATTIVA LA COLLEZIONE DI MACRO
    '//------------------------------------------------------------------------//

        Set objectMacros = CollectionMacro_PFunct()
        
    '//------------------------------------------------------------------------//


    ' STEP 01: RECUPERO PATH E FILE DI ORIGINE
    strSourceDbPath = PathFile_s_pFunct()
    If strSourceDbPath = "" Then
        MsgBox "ATTENZIONE: Path e nome del file di origine non specificati.", vbCritical, "Errore di importazione"
        Exit Sub
    End If

    ' STEP 02: APERTURA DEI DATABASE
    Set dbCurrent = CurrentDb
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(strSourceDbPath)

    ' STEP 03: RECUPERO DELLE MACRO DAL DATABASE DI ORIGINE
    strSQL = "SELECT Name FROM MSysObjects WHERE Type=-32766;" ' Tipo -32764 corrisponde alle macro
    Set qdf = dbSource.CreateQueryDef("", strSQL)
    Set rs = qdf.OpenRecordset()

    ' Loop attraverso le macro trovate nel database di origine
    Do While Not rs.EOF
        myNameMACRO_s = rs.Fields("Name").Value
        rs.MoveNext

        ' Controlla se la macro è nella collezione di macro
        objectExistsInCollectionMacros_Funct_b = InCollectionMacros_Funct_b(objectMacros, myNameMACRO_s)

        ' Se la macro esiste nella collezione, procedi con l'importazione
        If objectExistsInCollectionMacros_Funct_b Then
            ' Prova a eliminare la macro esistente nel database corrente
            On Error Resume Next
            
            DoCmd.DeleteObject acMacro, myNameMACRO_s
            
            deleteMACROCount_i = deleteMACROCount_i + 1
            deleteNameMACRO_s = deleteNameMACRO_s & myNameMACRO_s & vbCrLf
            On Error GoTo 0

            ' Importa la macro dal database di origine al database corrente
            DoCmd.TransferDatabase acImport, "Microsoft Access", strSourceDbPath, acMacro, myNameMACRO_s, myNameMACRO_s
            importedMACROesCount_i = importedMACROesCount_i + 1
            importedNameMACRO_s = importedNameMACRO_s & myNameMACRO_s & vbCrLf
        End If
    Loop

    ' Messaggio di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "Importazione completata:" & vbCrLf & _
               "Macro importate: " & importedMACROesCount_i & vbCrLf & _
               "Macro eliminate: " & deleteMACROCount_i & vbCrLf & _
               "Elenco macro eliminate: " & vbCrLf & deleteNameMACRO_s, vbInformation
    End If
    
    
                '//ciclo nella collezione per il controllo DELLE MACRO per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                     
                     '//itero nelle macro per la selezione dell'ultima
                    Set dbs = Application.CurrentProject
            
                     For Each obj In dbs.AllMacros
                     '//vado all'ultima query
                        DoCmd.SelectObject acMacro, obj.Name, True
                    
                     Next obj
                     
                     
            

                '//--------------------------------------------------------------------------------------//


    ' Libera le risorse
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set dbSource = Nothing
    Set dbCurrent = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Debug.Print "Errore durante l'importazione: " & Err.Description
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not dbSource Is Nothing Then Set dbSource = Nothing
    If Not dbCurrent Is Nothing Then Set dbCurrent = Nothing
End Sub

'//funzione di controllo Macro in collection
Function InCollectionMacros_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        If varItem = key Then
            InCollectionMacros_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionMacros_Funct_b = False
    On Error GoTo 0
End Function

'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** FINE ***
'//----------------------------------------------------------------------------------------//



'//=====================================================================================================================//
'//                            IMPORTO LE MACRO *** FINE ***
'//
'//=====================================================================================================================//
