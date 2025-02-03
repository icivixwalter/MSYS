

'//=========================================================================================//
'//
'//                 GESTIONE DEI MODULI ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@MODULI                    @GESTONE@MODULI
'//=========================================================================================//


'//IMPORTA GLI OGGETTI MODULI DEL DB ESTERNO  NEL DB CORRENTE *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti Table e MODULI controllando prima se _
            nel db corrente di destinazione già esiste la Table o MODULI per effettuare la cancellazione _
            prima dell'importazione.


'//PARAMETRI_DA_CAMBIARE _
        ----> Cmd_Importa_MODULI_Click _
        ----> Dim qdf As DAO.MODULIDef _
        ----> Dim strSourceDbPath As String _
        ----> Dim objectModules As Collection _
        ----> Dim objectExistsInCollectionModule_Funct_b As Boolean _
        ----> TABELLE E MODULI _
        ----> MODULI _
        ----> Dim objectExistsInCollectionModule_Funct_b As Boolean _
        ----> COLLECTION MODULI _
        ----> importedModulesCount_i _
        ----> deleteModulesCount_i _
        ----> deleteNameModule_s _

    '//I COMANDI BUTTON PER GESTIRE L'IMPORTAZIONI DEI MODULI: _
        Cmd_Importa_MODULI_Click         = ESEGUE IL CONTROLLO E L'IMPORTAZIONE _
        Cmd_Importa_MODULI_HELP_Click    = HELP GUIDA PER IL COMANDO DI IMPORTAZIONE MODULI

    '//OPERAZIONI ESEGUITE PER LA LAVORAZIONE DELLA COLLECTION MODULI _
        @MODULI@INSERT.01=CHIAMO LA FUNZIONE CHE CREA LA COLLECTION MODULI E LA CARICA CON GLI OGGETTI SCELTI _
        @MODULI@INSERT.02=PRECARICO LA COLLECTION MODULI DAIMPORTARE _
        @MODULI@INSERT.03=CREA LA COLLECTION MODULI _
        @MODULI@INSERT.04=SELECT TRA VARI OGGETTI MODULI _
        @MODULI@INSERT.05=RESET CASELLA DI TESTO _
        @MODULI@INSERT.06=STAMPO TESTATA DI RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI _
        @MODULI@INSERT.07=INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO _
        @MODULI@INSERT.08=ITERO NEGLI OGGETTI MODULI CARICATI E FACCIO UN STAMPA DI CONTROLLO E CARICO LA STRINGA DA STAMPARE _
        @MODULI@INSERT.09=STAMPO GLI OGGETTI NELLA CASELLA DI TESTO E LA AGGIORNO _
        @MODULI@INSERT.10=RESTITUISCO LA COLLEZIONE AL CHIAMANTE _
        @MODULI@INSERT.11=ITERO NELLA COLLECTION MODULI PER UNA STAMPA DI CONTROLLO DI STAMPA
     '//@MODULI@INSERT.12=Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO _
        @MODULI@INSERT.13=RECUPERO LA PATH ED IL FILE.MDB PER LE IMPORTAZIONI DALLE CASELLE DI TESTO _
        @MODULI@INSERT.14=IF MESSAGGIO DI IMPORTAZIONE : TRUE = MESSAGGIO FAMODULI NESSUN MESSAGGMODULI        @MODULI@INSERT.15=CONTROLLO SE PROGETTO INTERNO (NON CARICA i MODULI) OPPURE ESTERNO (CARICA i MODULI) _
        @MODULI@INSERT.17=INIZIO IMPORTAZIONE MODULI -> MA PRIMO CREO VARIABILI DB CORRENTE ED ESTERNO _
        @MODULI@INSERT.18=INIZIO IMPORTAZIONE MODULI -> APRO IL DB ESTERNO E CONFRONTO MODULI DB ESTERNO = MODULI COLLECTION _
        @MODULI@INSERT.19=INIZIO IMPORTAZIONE MODULI -> CHIAMO ROUTINE ESTERNA PER VERIFICARE ESISTENZA MODULI _
        @MODULI@INSERT.20=INIZIO IMPORTAZIONE MODULI -> ROUTINE ESTERNA VERIFICA SE MODULI DB ESTERNO ESISTE NELLA COLLETION MODULI DA CARICARE NEL  _DB CORRENTE _
        @MODULI@INSERT.21=INIZIO IMPORTAZIONE MODULI -> ROUTINE ESTERNA RESTITUISCE MODULI CHE GIA ESISTE NEL DB CORRENTE PRIMA CANCELLA POI IMPORTA (evito duplicatMODULI        @MODULI@INSERT.22=INIZIO IMPORTAZIONE MODULI -> CONTEGGIO i MODULI CANCEMODULI _
        @MODULI@INSERT.23=INIZIO IMPORTAZIONE MODULI -> CONTEGGIO i MODULI IMPORTATE _
        @MODULI@INSERT.23.BIS=TODO FARE IL CONTROLLO SULLA TABELLA DI SISTEMA ESTERNA?? _
        @MODULI@INSERT.23.TRIS=II° MESSAGGIO DI RIEPILOGO MODULI IMPORTATE E CANCELLATE ABILITATO=TRUE DISABILITATO=FALSE _
        @MODULI@INSERT.24=CICLO FOR PER ANDARE ALL'ULTIMA MODULI CARICATA E FARE IL DoCmd.SelectObject  DELL'OGGETTO _
        @MODULI@INSERT.25=FINE,PRIMA DI USCIRE RILASCIO GLI OGGETTI




'//Cmd_Importa_MODULI_Click         = ESEGUE IL CONTROLLO E L'IMPORTAZIONE
Private Sub Cmd_Importa_MODULI_Click()

Dim dbCurrent As DAO.Database
Dim dbSource As DAO.Database
Dim qdf As DAO.MODULIDef
Dim strSourceDbPath As String
Dim strObjectName As String
Dim objectModules As Collection
Dim varItem As Variant
Dim objectExistsInCollectionModule_Funct_b As Boolean

On Error GoTo ErrorHandler

'//=====================================================================================================================//
'//                            STEP_01 = RECUPERO PATH E COLLECTION
'//RECUPERO LA COLLEZIONE DEI MODULI DA IMPORTARE E FACCIO IL CONTROLLO
'//=====================================================================================================================//

    '//RESET
    'Inizializza i contatori e le stringhe di riepilogo
    iCount = 0
    importedModulesCount_i = 0
    importedObjects_s = ""
    deleteModulesCount_i = 0
    deleteNameModule_s = ""
    
    
    '//@MODULI@INSERT.01=CHIAMO LA FUNZIONE CHE CREA LA COLLECTION MODULI E LA CARICA CON GLI OGGETTI SCELTI
    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION MODULI per ottenere la collezione di TABELLE E MODULI
    Set objectModules = CollectionModuli_PFunct()
    
            '//2024.08.03_CollectionModuli_PFunct = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE _
                        faccio un controllo sul numero degli oggetti della collection se = 0 allora _
                        non ci sono oggetti da caricare per cui si esce dalla routine.
                        
            If objectModules.count = 0 Then Exit Sub
                
    '//------------------------------------------------------------------------//
    
    
    
    '//@MODULI@INSERT.11=ITERNO NELLA COLLECTION MODULI PER UNA STAMPA DI CONTROLLO DI STAMPA
    '//I° CONTROLLO SULLA COLLECTION MODULI RECUPERATA
    '//ciclo nella collezione per il controllo delle TABELLE E MODULI precaricate
    For Each objCollection_objectName_v In objectModules
        Debug.Print "Oggetto nella collezione: " & objCollection_objectName_v
        '//conteggio oggetti DA IMPORTARE
        iCount = iCount + 1
        importedNameModules_s = importedNameModules_s & objCollection_objectName_v & vbCrLf
    Next objCollection_objectName_v
    
    
    '//@MODULI@INSERT.12=Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO OGGETTI COLLECTION PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Oggetti COLLECTION PRECARICATA DA IMPORTARE: " & iCount & vbCrLf & importedNameModules_s & _
               "Oggetti esistenti cancellati: " & deleteModulesCount_i & vbCrLf & deleteNameModule_s, vbInformation
    End If
    
            '//RESET
            iCount = 0
             
    

'//---------------------------------------------------------------------------//


'//@MODULI@INSERT.13=RECUPERO LA PATH ED IL FILE.MDB PER LE IMPORTAZIONI DALLE CASELLE DI TESTO
'//---------------------------------------------------------------------------//
'//NOTE: chiamo la funzione e recupero la path ed il file.mdb facendo un controllo _
        se la variabile è nulla msg di avviso ed uscita dalla routine.
    '//reset
    sourceDBPath_s = ""

    '//@RECUPERO@PATH_(ed il file
    sourceDBPath_s = PathFile_s_pFunct

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
'//                             STEP_02 = @CONTROLLO@PROGETTO
'//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
'//=====================================================================================================================//
'//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
         non si possono caricare gli oggetti nel PROGETTO MSYS_N01_OGGETTI.mdb perhè è originale _
         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
         per caricare gli oggetti table, MODULI , form ecc. da importare. _
         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
            
            
           '//@MODULI@INSERT.14=IF MESSAGGIO DI IMPORTAZIONE : TRUE = MESSAGGIO FALSO = NESSUN MESSAGGIO
           '//01) PRIMO CONTROLLO _
             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
           '//..........................................................//
             '//Attenzione L'AVVISO produce il messaggio solo se il _
                parametro è TRUE altrimenti il messaggio viene escluso.
             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
             
             If MyValue = 0 Then Exit Sub
           '//..........................................................//
                 
             MODULI           '//@MODULI@INSERT.15=CONTROLLO SE PROGETTO INTERNO (NON CARICA i MODULI) OPPURE ESTERNO (CARICA i MODULI)
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
'//                             STEP_02 = @CONTROLLO@PROGETTO
'//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
'//=====================================================================================================================//

'//===================================================================================================================MODULI'//                            STEP_03 = IMPORTO LE TABELLE E i MODULI
'//
'//=====================================================================================================================//
'//@MODULI@INSERT.16= INIZIO IMPORTAZIONE MODULI

    '//RESET
    '//Inizializza i contatori e le stringhe di riepilogo
    
    
    importedQueriesCount_i = 0
    deleteModulesCount_i = 0
    deleteNameModule_s = ""
    importedNameModules_s = ""
    
    
    '//@MODULI@INSERT.17=INIZIO IMPORTAZIONE MODULI -> MA PRIMO CREO VARIABILI DB CORRENTE ED ESTERNO
    ' Apri il database corrente
    Set dbCurrent = CurrentDb

    ' Apri il database di origine
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath_s)
    
    '//@MODULI@INSERT.18=INIZIO IMPORTAZIONE MODULI -> APRO IL DB ESTERNO E CONFRONTO MODULI DB ESTERNO = MODULI COLLECTION
    'Loop attraverso tutte i MODULI nel database di origine ossia del DATABASE ESTERNO
    For Each qdf In dbSource.MODULIDefs
        ' Salta le tabelle di sistema (i cui nomi iniziano con "MSys") aggiungo come parametro( la collezione table + nome tabella)
        If Left(Mdl.Name, 3) <> "~sq" Or (Left(Mdl.Name, 3) = "~sq" And InCollectionModule_Funct_b(objectModules, Mdl.Name)) Then
            ' Nome della tabella nel database di origine
            myNameModuli_s = Mdl.Name
            
            Debug.Print
            Debug.Print "CONTROLLO MODULI DA IMPORTARE --> : "; myNameModuli_s
            Debug.Print
            
            
            ' Controlla se la MODULI è nella collezione
            objectExistsInCollectionModule_Funct_b = False
            
            For Each varItem In objectModules
                If myNameModuli_s = varItem Then
                    objectExistsInCollectionModule_Funct_b = True
                    Exit For
                End If
            Next varItem
            
            '//@MODULI@INSERT.19=INIZIO IMPORTAZIONE MODULI -> CHIAMO ROUTINE ESTERNA PER VERIFICARE ESISTENZA MODULI
            ' Se la MODULI è nella collezione, importala
            If objectExistsInCollectionModule_Funct_b Then
            
                '//@MODULI@INSERT.21=INIZIO IMPORTAZIONE MODULI -> ROUTINE ESTERNA RESTITUISCE MODULI CHE GIA ESISTE NEL DB CORRENTE PRIMA CANCELLA POI IMPORTA (evito duplicati)
                ' Elimina la MODULI se esiste già nel database corrente
                On Error Resume Next
                dbCurrent.MODULIDefs.Delete myNameModuli_s
            MODULI                '//@MODULI@INSERT.22=INIZIO IMPORTAZIONE MODULI -> CONTEGGIO i MODULI CANCELLATE
                deleteModulesCount_i = deleteModulesCount_i + 1
                deleteNameModule_s = deleteNameModule_s & myNameModuli_s & vbCrLf
                
                On Error GoTo 0
                
                ' Copia la tabella dal database di origine al database corrente
                DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acMODULI, myNameModuli_s, myNameModuli_s
                
                '//conteggio TABELLE IMPORMODULI                '//@MODULI@INSERT.23=INIZIO IMPORTAZIONE MODULI -> CONTEGGIO i MODULI IMPORTATE
                importedModulesCount_i = importedModulesCount_i + 1
                importedNameModules_s = importedNameModules_s & myNameModuli_s & vbCrLf
                
            End If
        End If
    Next qdf
                
        '//@MODULI@INSERT.23.BIS=TODO FARE IL CONTROLLO SULLA TABELLA DI SISTEMA ESTERNA?? _
        '//TODO: MANCA IL CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects DA CREARE VEDI FUNZIONE Cmd_Importa_MODULI_Click
        '//----------------------------------------------------------------------------------------------------------------------//
            
        
            '//TODO: adattare il controllo sulla tabella di sistema esterna ... vai a Cmd_Importa_MODULI_Click

        '//----------------------------------------------------------------------------------------------------------------------//
        
    '//@MODULI@INSERT.23.TRIS=II° MESSAGGIO DI RIEPILOGO MODULI IMPORTATE E CANCELLATE ABILITATO=TRUE DISABILITATO=FALSE _
    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO OGGETTI PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Oggetti importati: " & importedModulesCount_i & vbCrLf & importedNameModules_s & vbCrLf & _
               "Oggetti esistenti cancellati: " & deleteModulesCount_i & vbCrLf & deleteNameModule_s, vbInformation
    End If
                
                '//@MODULI@INSERT.24=CICLO FOR PER ANDARE ALL'ULTIMA MODULI CARICATA E FARE IL DoCmd.SelectObject  DELL'OGGETTO
                '//ciclo nella collezione per il controllo deI MODULI precaricate per SELEZIONARE LA PRIMA
                '//--------------------------------------------------------------------------------------//
                    For Each objCollection_objectName_v In objectModules
                        
                        
                        '//vado all'ultima MODULI
                        DoCmd.SelectObject acMODULI, objCollection_objectName_v, True
                        
                    Next objCollection_objectName_v
                                
                    '//@MODULI@INSERT.25=FINE,PRIMA DI USCIRE RILASCIO GLI OGGETTI
                    Set dbCurrent = Nothing
                    Set objectModules = Nothing
                    
                '//--------------------------------------------------------------------------------------//

'//===================================================================================================================MODULI'//                            STEP_03 = IMPORTO LE TABELLE E i MODULI *** fine ***
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
End Sub   '//Private Sub Cmd_Importa_MODULI_Click()

'//@MODULI@INSERT.20=INIZIO IMPORTAZIONE MODULI -> ROUTINE ESTERNA VERIFICA SE MODULI DB ESTERNO ESISTE NELLA COLLETION MODULI DA CARICARE NEL DB CORRENTE
' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionModule_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    For Each varItem In col
        If varItem = key Then
            InCollectionModule_Funct_b = True
            Exit Function
        End If
    Next varItem
    InCollectionModule_Funct_b = False
    On Error GoTo 0
End Function

'//IMPORTA GLI OGGETTI MODULI DEL DB ESTERNO  NEL DB CORRENTE  *** FINE ***
'//----------------------------------------------------------------------------------------//



'//IMPORTA MODULI HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//
'//Cmd_Importa_MODULI_HELP_Click    = HELP GUIDA PER IL COMANDO DI IMPORTAZIONE MODULI
Private Sub Cmd_Importa_MODULI_HELP_Click()

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
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_IMPORTA_MODULI.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA MODULI HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=========================================================================================//
'//
'//                 GESTIONE DEI MODULI ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@MODULI
'//=========================================================================================//


