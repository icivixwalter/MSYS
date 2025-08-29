# Modello_IMPORTA_OGGETTI_QUERY.md


'//IMPORTA GLI OGGETTI TABLE E QUERY DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti Table e Query controllando prima se _
            nel db corrente di destinazione già esiste la Table o Query per effettuare la cancellazione _
            prima dell'importazione.


'//PARAMETRI_DA_CAMBIARE _ 
		----> Cmd_Importa_QUERY_Click _
		----> Dim qdf As DAO.QueryDef _
		----> Dim strSourceDbPath As String _
		----> Dim objectQueries As Collection _
		----> Dim objectExistsInCollectionQueries_Funct_b As Boolean _
		----> TABELLE E QUERY _
		----> QUERY _
		----> Dim objectExistsInCollectionQueries_Funct_b As Boolean _
		----> COLLECTION QUERY _
		----> importedQueryesCount_i _
		----> deleteQueryesCount_i _
		----> deleteNameQueryes_s _

Private Sub Cmd_Importa_QUERY_Click()

Dim dbCurrent As DAO.Database
Dim dbSource As DAO.Database
Dim qdf As DAO.QueryDef
Dim strSourceDbPath As String
Dim strObjectName As String
Dim objectQueries As Collection
Dim varItem As Variant
Dim objectExistsInCollectionQueries_Funct_b As Boolean

On Error GoTo ErrorHandler

'//=====================================================================================================================//
'//                            STEP_01 = RECUPERO PATH E COLLECTION
'//RECUPERO LA COLLEZIONE TABELLE E QUERY DA IMPORTARE E FACCIO IL CONTROLLO
'//=====================================================================================================================//

    '//RESET
    'Inizializza i contatori e le stringhe di riepilogo
    icount = 0
    importedQueryesCount_i = 0
    importedObjects_s = ""
    deleteQueryesCount_i = 0
    deleteNameQueryes_s = ""
              
    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION QUERY per ottenere la collezione di TABELLE E QUERY
    Set objectQueries = CollectionQuery_PFunct()
    '//------------------------------------------------------------------------//

    '//I° CONTROLLO SULLA COLLECTION QUERY RECUPERATA
    '//ciclo nella collezione per il controllo delle TABELLE E QUERY precaricate
    For Each objCollection_objectName_v In objectQueries
        Debug.Print "Oggetto nella collezione: " & objCollection_objectName_v
        '//conteggio oggetti DA IMPORTARE
        icount = icount + 1
        importedNameQueryes_s = importedNameQueryes_s & objCollection_objectName_v & vbCrLf
    Next objCollection_objectName_v

    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO OGGETTI COLLECTION PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Oggetti COLLECTION PRECARICATA DA IMPORTARE: " & icount & vbCrLf & importedNameQueryes_s & _
               "Oggetti esistenti cancellati: " & deleteQueryesCount_i & vbCrLf & deleteNameQueryes_s, vbInformation
    End If
    
            '//RESET
            icount = 0
             
    

'//---------------------------------------------------------------------------//

'//RECUPERO LA PATH ED IL FILE.MDB PER LE IMPORTAZIONI
'//---------------------------------------------------------------------------//
'//NOTE: cheiamo la funzione e recupero la path ed il file.mdb facendo un controllo _
        se la variabile è nulla msg di avviso ed uscita dalla routine.
    '//reset
    sourceDBPath_s = ""

    '//@RECUPERO@PATH_(ed il file
    sourceDBPath_s = PathtFile_s_pFunct

    If sourceDBPath_s = "" Then
        MsgBox "ATTENZIONE ERRORE path e file sono null uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        If MyValue = 0 Then Exit Sub
    End If

      ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO GLI OGGETTI DA QUESTO DATABASE; path e del file : " & sourceDBPath_s, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//=====================================================================================================================//
'//                            STEP_01 = RECUPERO PATH E COLLECTION
'//RECUPERO LA COLLEZIONE FORM DA IMPORTARE E FACCIO IL CONTROLLO          **** FINE ***
'//=====================================================================================================================//

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

'//=====================================================================================================================//
'//                            STEP_03 = IMPORTO LE TABELLE E LE QUERY
'//
'//=====================================================================================================================//
    
    '//RESET
    '//Inizializza i contatori e le stringhe di riepilogo
    
    
    importedQueriesCount_i = 0
    deleteQueryesCount_i = 0
    deleteNameQueryes_s = ""
    importedNameQueryes_s = ""
    

    ' Apri il database corrente
    Set dbCurrent = CurrentDb

    ' Apri il database di origine
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath_s)
    
    ' Loop attraverso tutte le tabelle nel database di origine
    For Each qdf In dbSource.QueryDefs
        ' Salta le tabelle di sistema (i cui nomi iniziano con "MSys") aggiungo come parametro( la collezione table + nome tabella)
        
        If Left(qdf.Name, 3) <> "~sq" Or (Left(qdf.Name, 3) = "~sq" And InCollectionQueries_Funct_b(objectQueries, qdf.Name)) Then
            ' Nome della tabella nel database di origine
            myNameQuery_s = qdf.Name
            
            ' Controlla se la tabella è nella collezione
            objectExistsInCollectionQueries_Funct_b = False
            
            For Each varItem In objectQueries
                If myNameQuery_s = varItem Then
                    objectExistsInCollectionQueries_Funct_b = True
                    Exit For
                End If
            Next varItem
            
            ' Se la tabella è nella collezione, importala
            If objectExistsInCollectionQueries_Funct_b Then
                ' Elimina la tabella se esiste già nel database corrente
                On Error Resume Next
                
                dbCurrent.QueryDefs.Delete myNameQuery_s
                deleteQueryesCount_i = deleteQueryesCount_i + 1
                deleteNameQueryes_s = deleteNameQueryes_s & myNameQuery_s & vbCrLf
                
                On Error GoTo 0
                
                ' Copia la tabella dal database di origine al database corrente
                DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acQuery, myNameQuery_s, myNameQuery_s
                
                '//conteggio TABELLE IMPORTATE
                
                importedQueryesCount_i = importedQueryesCount_i + 1
                importedNameQueryes_s = importedNameQueryes_s & myNameQuery_s & vbCrLf
                
            End If
        End If
    Next qdf
                
        '//TODO: MANCA IL CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects DA CREARE VEDI FUNZIONE Cmd_Importa_MODULI_Click
        '//----------------------------------------------------------------------------------------------------------------------//
        
            '//TODO: adattare il controllo sulla tabella di sistema esterna ... vai a Cmd_Importa_MODULI_Click

        '//----------------------------------------------------------------------------------------------------------------------//
                
    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO OGGETTI PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Oggetti importati: " & importedQueryesCount_i & vbCrLf & importedNameQueryes_s & vbCrLf & _
               "Oggetti esistenti cancellati: " & deleteQueryesCount_i & vbCrLf & deleteNameQueryes_s, vbInformation
    End If
    
                '//ciclo nella collezione per il controllo delle QUERY precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                    For Each objCollection_objectName_v In objectQueries
                        
                        '//vado all'ultima query
                        DoCmd.SelectObject acQuery, objCollection_objectName_v, True
                        
                    Next objCollection_objectName_v
                                
                    '//PRIMA DI USCIRE RILASCIO GLI OGGETTI
                    Set dbCurrent = Nothing
                    Set objectQueries = Nothing
                    
                '//--------------------------------------------------------------------------------------//

'//=====================================================================================================================//
'//                            STEP_03 = IMPORTO LE TABELLE E LE QUERY *** fine ***
'//
'//=====================================================================================================================//

'// CHIUSURA E CONTROLLO ERRORI
'//------------------------------------------------------------------------------------------------//
Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Debug.Print "Errore durante l'importazione: " & Err.Description, vbCritical
    If Not dbSource Is Nothing Then
        dbSource.Close
    End If
    Set dbSource = Nothing
    Set dbCurrent = Nothing
End Sub   '//Private Sub Cmd_Importa_QUERY_Click()


' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionQueries_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        If varItem = key Then
            InCollectionQueries_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionQueries_Funct_b = False
    On Error GoTo 0
End Function

'//IMPORTA GLI OGGETTI TABLE E QUERY DEL DB ESTERNO IN TABELLA  *** FINE ***
'//----------------------------------------------------------------------------------------//

