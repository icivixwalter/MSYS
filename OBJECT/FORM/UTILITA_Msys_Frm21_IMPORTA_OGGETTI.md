UTILITA_Msys_Frm21_IMPORTA_OGGETTI.md
   @form@UTILITA_(classe della form utilita)

@INIZIO

### CODICE_CLASSE
'//DELETE QUERY HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=====================================================================================================================//
'//                     DELETE OGGETTI QUERY DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE QUERY HELP
'//=====================================================================================================================//







'//************************************************************************************************************************//
'//***
'//***                              COLLECTION OBJECT   *** FINE ***
'//************************************************************************************************************************//



'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI FORM DEL DB ESTERNO *** INIZIO ***
'//                     @IMPORTA@OGGETTI@FORM@MSYS
'//=====================================================================================================================//



'//IMPORTA TUTTI GLI OGGETTI FORM DEL DB MSYSDBEST - LE ATTIVITA
'//----------------------------------------------------------------------------------------//
'//@01_@IMPORTA@FORM@DIM                  -------------------> DIM VARIABILI
'//@02_@IMPORTA@FORM@REST                 -------------------> RESET DELLE VARIABILI
'//@03_@IMPORTA@FORM@LOG                  -------------------> LOG FILE INIZIALE
'//@04_@IMPORTA@FORM@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
'//@05_@IMPORTA@@FORMCONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
'//@06_@IMPORTA@FORM@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
'//@07_@IMPORTA@FORM@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
'//@08_@IMPORTA@FORM@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
'//@09_@IMPORTA@FORM@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
'//@10_@IMPORTA@FORM@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
'//@11_@IMPORTA@FORM@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
'//@12_@IMPORTA@FORM@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
'//@13_@IMPORTA@FORM@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
'//@14_@IMPORTA@FORM@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
'//@15_@IMPORTA@FORM@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
'//@15.BIS_@IMPORTA@FORM@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
'//@16_@IMPORTA@FORM@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
'//@17_@IMPORTA@FORM@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
'//@18_@IMPORTA@FORM@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI


Private Sub Cmd_Importa_FORM_Click()
    
    '//@01_@IMPORTA@FORM@DIM                  -------------------> DIM VARIABILI
    ' Dichiarazione delle variabili
    '//oggetti form e variabile nome form
    Dim formsCollection As Collection                       '//oggetti collezione
    Dim formName_v As Variant
    Dim deletedForms_s As String
    Dim importedForms_s As String
    Dim myformName_s As String
    
    '//il database + path
    Dim sourceDB As DAO.Database
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDBPath As String
    
    '//contatori oggetti: TOTALE OGGETTI PREVISTI; da importare e cancellare
    Dim TotCollectioForms_i As Integer                          '//totale degli oggetti form della collezione da importare
    Dim importedCount_i As Integer                              '//oggetti importati
    Dim deletedCount_i As Integer                               '//oggetti cancellati
    
    
    '//bool e log file
    Dim Bool1 As Boolean
    Dim MyValue As Integer
    Dim logFile As Integer
    
    On Error GoTo ErrorHandler
    
    '//@02_@IMPORTA@FORM@REST                 -------------------> RESET DELLE VARIABILI
    '//RESET = Inizializza i contatori e le stringhe di riepilogo
    TotCollectioForms_i = 0
    importedCount_i = 0
    deletedCount_i = 0
    deletedForms_s = ""
    importedForms_s = ""
    TotPerc_dbl = 0
    ProgressPercentTwips_dbl = 0
    
    '//@03_@IMPORTA@FORM@LOG                  -------------------> LOG FILE INIZIALE
    '@FILE@LOG@FORM: Apre il file di log per scrivere i dettagli dell'importazione apre nella path corrente.
    logFile = FreeFile()
    Open CurrentProject.PATH & "\LOG_IMPORTA_FORM.TXT" For Output As #logFile
    Print #logFile, "Log Importazione Form - " & Now
    Print #logFile, "------------------------------------------------"
    Print #logFile,
    
    '//@04_@IMPORTA@FORM@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
    'Recupero la collezione di form da importare
      Set formsCollection = CollectionForm_PFunct()
      
    '//@05_@IMPORTA@@FORMCONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
    '//se non esistono oggetti collection EXIT
    If formsCollection.Count = 0 Then
        Print #logFile, "Nessuna form da importare."
        Close #logFile
        '//Prima di uscire restto l'oggetto collection
        Set formsCollection = Nothing
        Exit Sub
    End If
    
    '//@06_@IMPORTA@FORM@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
    '//RECUPERO LA PATH =  Controllo se il percorso del database di origine è valido
    sourceDBPath = PathFile_s_pFunct
    
    '//@07_@IMPORTA@FORM@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
    If sourceDBPath = "" Then
        MsgBox "Errore: path e file nulli.", vbCritical
        Print #logFile, "Errore: path e file nulli."
        Close #logFile
        Exit Sub
    End If
    
    '//@08_@IMPORTA@FORM@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
    '//SE NON ESISTE IL FILE DB = Controllo se il file del database di origine esiste
    If Dir(sourceDBPath) = "" Then
        MsgBox "Errore: il file del database di origine non esiste.", vbCritical
        Print #logFile, "Errore: il file del database di origine non esiste."
        Close #logFile
        Exit Sub
    End If
    
    '//@09_@IMPORTA@FORM@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
    '//.......................................................................................//
      '//AVVISI MESSAGGIO IMPOSTATO = TRUE - >  Controllo del valore di avviso
      MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
      If MyValue = 0 Then Exit Sub
      
      '//CONTROLLO NOME PROGETTO =  Controllo del nome del progetto
      Bool1 = CheckProjectName_b_PFunct()
      If Bool1 = True Then Exit Sub
      
    '//.......................................................................................//
    
    
        '//@10_@IMPORTA@FORM@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI  @progress@bar
        '//********************************************************************************************************************//
        '//@PROGRES@BAR@FASE_01_(imposto le variabili e recupero i valori dalle funzioni)
        
                
                '//NOTE: devi creare 4 funzioni per avere i 4 parametri per chiamare la progress bar: _
                     @01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                     @02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl _
                     @03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     @04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     @05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                     
             
                  
                      '//@01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                             Imposto il totale oggetti form da importare se esistono
                      TotCollectioForms_i = formsCollection.Count
             
                    
                    '//@02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl
                       Twips_dbl = Calcola_Twips_pFunct(TotCollectioForms_i)
                    
                    '//@03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     '//chiama la funzione percentuale incremento e ne restituisce la _
                         percentuale come double
                     PercIncr_dbl = PercentualeIncremento_pFunct(TotCollectioForms_i)             '//calcolo il numero di collection form da controllare
         
                     '//@04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     '//chiama la funzione TIPO IMPORT
                         TipoImport_s = TipoImport_pFunct("\LOG_IMPORTA_FORM.TXT")
         
                     
                     '//@05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                         '//Apri la form della progressBar
                         '//
                             Call ApriProgressBar
                             ' Aspetta che il form sia effettivamente caricato
                              DoEvents
                         
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI *** FINE ***
        '//********************************************************************************************************************//
    
    
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** INIZIO ****
    '//===========================================================================================================================//
            
          '//@11_@IMPORTA@FORM@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
          ' Se siamo su un progetto esterno (Bool1 = False), procedo con l'importazione
          If Bool1 = False Then
                ' Apertura del database di origine e connessione al database corrente
                Set ws = DBEngine.Workspaces(0)
                Set sourceDB = ws.OpenDatabase(sourceDBPath)
                Set db = CurrentDb
                
                
                '//reset
                ProgressPercentTwips_dbl = 0
                
                
                '//@12_@IMPORTA@FORM@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
                ' Ciclo attraverso la collezione di form da importare
                For Each formName_v In formsCollection
                    myformName_s = formName_v
                    
                    '//@13_@IMPORTA@FORM@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
                    ' Controllo se la form esiste nel database di origine
                    If ObjectExists(sourceDB, acForm, myformName_s) Then
                        ' Se la form esiste nel database corrente, la elimino prima di importarla
                        If ObjectExists(db, acForm, myformName_s) Then
                            DoCmd.DeleteObject acForm, myformName_s
                            deletedCount_i = deletedCount_i + 1
                            deletedForms_s = deletedForms_s & myformName_s & vbCrLf
                            Print #logFile, "Form eliminata: " & myformName_s
                             
                              
                        End If
                        
                          '//@14_@IMPORTA@FORM@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
                          ' Importazione della form dal database di origine
                          DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acForm, myformName_s, myformName_s
                          importedCount_i = importedCount_i + 1
                          importedForms_s = importedForms_s & myformName_s & vbCrLf
                          
                          '//SCRIVO SUL @FILE@LOG
                          Print #logFile, "NRO: " & importedCount_i & " Form importata: " & myformName_s
                          
                              
                        '//@15_@IMPORTA@FORM@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   @progress@bar
                        '//********************************************************************************************************************//
                        '//NOTE: impostazioni e recupero valori dalla funzioni base per la progress bar.
                            
                              '// RECUPERO il messaggio tipo di importazione delle collection e numero di oggetti da importare .
                                TipoImport_s = "FORM IMPORTATE NRO : " & importedCount_i & " su TOTALE OGGETTI: " & TotCollectioForms_i       '// STRINGA MESSAGGIO PER IL TIPO DI IMPORTAZIONE E NRO OGGETTI
                                
                                '//LA PERCENTUALE INCREMENTATA = totale di incremento % per ogni oggetto importato
                                TotPerc_dbl = TotPerc_dbl + PercIncr_dbl
                                
                                '//TOTALE % DI INCREMENTO IN TWIPS, con la ProgressPercentTwips_dbl è la variabile che indica la %di incremento per ogni oggetto + i _
                                   twips necessari per ogni elemento (su 10 totale della barra)
                                ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + (PercIncr_dbl + Twips_dbl)          '//la % di accrescimento barra
                                
                                '//TODO: DA CANCELLARE...
                                'Dbl1 = Dbl1 + Twips_dbl                                         '//twips di aumento barra
                                
                                
                                          '//@15.BIS_@IMPORTA@FORM@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
                                          '//CHIAMO LA FUNZIONE @AGGIORNA@PROGRESS@BAR
                                          '//.................................................................//
                                          '//note: chiamo la routine di aggiornamento della progress bar che _
                                              puo essere utilizzata per ulteriori integrazioni di calcolo. _
                                              Devo inviare ma devo inviare 5 parametri : _
                                                  TotCollectioForms_i       = il TOTALE POTENZIALE degli oggetti da importare _
                                                  NroCollection_i           = il numero della collezione per il calcolo; _
                                                  Twips                     = i twip di incremento _
                                                  ProgressPercentTwips_dbl  =  La percentuale di incremento ( ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + PercIncr_dbl) _
                                                  TipoImport_s              = il tipo di importazione
                                                  
                                              AggiornaProgressBar _
                                                          TotCollectioForms_i, _
                                                          ProgressPercentTwips_dbl, _
                                                          Twips_dbl, _
                                                          TotPerc_dbl, _
                                                          TipoImport_s
                                          '//.................................................................//
        
                               
                                T = Timer
                                While Timer < T + 0.001
                                    DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                Wend
                          
                                  
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   *** FINE ***
                        '//********************************************************************************************************************//
        
            
                        
                    Else
                        
                        '//@16_@IMPORTA@FORM@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
                        ' Se la form non esiste nel database di origine, lo scrivo nel log
                        Print #logFile, "Form non trovata nel database di origine: " & myformName_s
                    End If
                
                Next formName_v
        
        
                                        
                        '//@17_@IMPORTA@FORM@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
                        '//CHIUSURA CON RITARDO DELLA BARRA DI PROGRESSIONE @progress@bar
                        '//---------------------------------------------------------------------------//
                    
                                        T = Timer
                                        
                                        '// RITARDO CHIUSURA DI 5 SECONDI
                                        While Timer < T + 5
                                            DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                        Wend
                                        
                                        ' Chiudi la form frmProgressBar
                                        DoCmd.Close acForm, "frmProgressBar", acSaveNo
                                        
                        '//---------------------------------------------------------------------------//
                        
                        
                        
                '//@18_@IMPORTA@FORM@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI
                ' Chiudo il database di origine
                  Set frm_obj = Nothing                 ' Rilascia il riferimento al form PROGRES BARR
                
                '//chiudo il db
                sourceDB.Close
                Set sourceDB = Nothing
                Set ws = Nothing
            
                ' Messaggio finale con riepilogo delle operazioni
                If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                    MsgBox "Importazione completata." & vbCrLf & _
                        "Forms importate: " & importedCount_i & vbCrLf & importedForms_s & _
                        "Forms esistenti cancellate: " & deletedCount_i & vbCrLf & deletedForms_s, vbInformation
                End If
            
                ' Scrivo nel file di log il riepilogo finale
                Print #logFile, "Importazione completata con " & importedCount_i & " forms importate."
                Close #logFile
            
                ' Pulizia delle variabili del db e della collection forms
                Set db = Nothing
                Set formsCollection = Nothing
                
          End If '//If Bool1 = False Then = Progetto esterno
      
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** FINE ****
    '//===========================================================================================================================//
    
    
'// USCITA E GESTIONE ERRORI
'//---------------------------------------------------------------------------------------------------------//

Exit Sub
    
ErrorHandler:
    ' Gestione degli errori con messaggio a video e scrittura nel log
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Print #logFile, "Errore durante l'importazione: " & Err.Description
    If Not sourceDB Is Nothing Then sourceDB.Close
    Close #logFile
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub



























