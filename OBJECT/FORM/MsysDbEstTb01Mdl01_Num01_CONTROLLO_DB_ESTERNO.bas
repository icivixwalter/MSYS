Attribute VB_Name = "MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO"
'//MODULO = MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO _
            MODULO PER IL CONTROLLO DELLE TABELLE DI SISTEMA.


'//ACCESSO AL DB ESTERNO - controllo tabelle
'//============================================================================================//
'//NOTE : la routine e la funzione aprono una istanza presso il db esterno access per il _
        controllo delle tabelle. La routine utilizza 3 collezioni di oggetti che per quanto _
        riguarda la systemTables = precarico quali tabelle sono di sistema poi crea altre _
        due collection che vengono popolate se il controllo if IsSystemTable restituisce false _
        perche chiama la funzione per il controllo se la tabella appartiene al sistema, per esclusione _
        appartiene alla fisiche o alle collegate, e quindi popola le due collection che poi vengono _
        stampate
    

Sub ListTablesInExternalDB()
    '//DIM path, db e tabella
    Dim dbPath As String
    Dim externalDB As DAO.Database
    Dim tdf As DAO.TableDef
    '//dim le collection per separare il tipo di tabelle
    Dim systemTables As Collection
    Dim connectedTables As Collection
    Dim physicalTables As Collection
    
    ' Percorso del database esterno
    dbPath = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU_TB03_OGGETTI_DA_CANCELLARE\MENU_TB03_OGGETTI_DA_CANCELLARE.mdb"
    
    '//popola la collection delle tabelle di sistema
    ' Aggiungi i nomi delle tabelle di sistema da escludere
    Set systemTables = New Collection
    systemTables.Add "MSysAccessObjects"
    systemTables.Add "MSysACEs"
    systemTables.Add "MSysObjects"
    systemTables.Add "MSysQueries"
    systemTables.Add "MSysAccessXML"
    systemTables.Add "MSysRelationships"
    systemTables.Add "MSysNavPaneGroupCategories"
    systemTables.Add "MSysNavPaneGroupToObjects"
    systemTables.Add "MSysNavPaneObjectIDs"
    systemTables.Add "MSysNavPaneGroups"

    ' Inizializza le collezioni per tabelle fisiche e collegate
    Set connectedTables = New Collection
    Set physicalTables = New Collection

    ' Apri il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(dbPath)
    
    
    
       
                
                        
                    ' Scansiona tutte le tabelle nel database esterno
                    '//...............................................................................//
                        For Each tdf In externalDB.TableDefs
                            ' Verifica se la tabella è di sistema con il confronte nella collection precaricata _
                              e chiama funzione di controllo
                            If IsSystemTable(tdf.Name, systemTables) Then
                                Debug.Print tdf.Name & " (Tavola di sistema)"
                                
                                'se la tabella è collegata aggiunge alla collection
                            ElseIf Len(tdf.Connect) > 0 Then
                                connectedTables.Add tdf.Name
                            Else
                                
                                'la tabella è fisica e la aggiunge alla collection
                                physicalTables.Add tdf.Name
                            End If
                        
                        Next tdf
                                
                                    ' Stampa le tabelle collegate con le collezioni
                                    Debug.Print
                                    Debug.Print
                                    Debug.Print "Tabelle collegate:"
                                    'iterazione nella collection delle tabelle collegate
                                    For Each t In connectedTables
                                        Debug.Print t & " (Collegata)"
                                    Next t
                                
                                    ' Stampa le tabelle fisiche
                                    Debug.Print
                                    Debug.Print
                                    Debug.Print "Tabelle fisiche:"
                                    '//iterazione nella collezione fisica
                                    For Each t In physicalTables
                                        Debug.Print t & " (Fisica)"
                                    Next t
                                
                    
                    '//...............................................................................//
                        

            
            

     
    
    ' Chiudi il database esterno
    externalDB.Close
    Set externalDB = Nothing
End Sub

'//FUNZIONE PER IL CONTROLLO DELLE TABELLE DI SISTEMA
'//controlla la tabella corrente con l'elenco della collection systemTables
Function IsSystemTable(tableName As String, systemTables As Collection) As Boolean
    Dim tbl As Variant
    IsSystemTable = False
    For Each tbl In systemTables
        If tableName = tbl Then
            IsSystemTable = True
            Exit For
        End If
    Next tbl
End Function


'//============================================================================================//

