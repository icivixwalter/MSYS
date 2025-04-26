VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_CLASSE_GESTIONE_FILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database



'/
'//****************************************************************************************//'
'//**                   CLASSE_TB01_GESTIONE   CLASSE TB01                               **// '
'//**                                                                                    **// '
'//****************************************************************************************//'
'//**      @classe@gestione_(apertura file)_@WALTER                                      **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**      @APERTURA@DEL@FILE@UTILIZZO DELLA CLASSE PER L'APERTURA DEL FILE              **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//**                                                                                    **// '
'//****************************************************************************************//'
      
    
    '//@APERTURA@DEL@FILE      @UTILIZZO@DELLA@CLASSE_(per gestire l'apertura del file vedi questo modello di apertura)
    '//.....................................................................................................
    '//NOte: @MODELLO@APERTURA@FILE_(QUESTO E' IL MODELLO PER APRIRE UN FILE UTILIZZANDO LA FUNZIONE DELLA CLASSE) _
              MODELLO ISTANZIAMENTO CLASSE ATTI DI GESTIONE E @APRI@FILE_(APRO IL FILE UTILIZZANDO LA CLASSE FILE)
        
        
        '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
            'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione
    
        '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
           l'oggetto di classe alla variabile e creo creo l'oggetto di classe
            'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione
        
        '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
            'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
            'MyCls_CLASSE_ATTI_Tb01_Gestione.APRO_PATH_O_FILE_GENERICI_pSub (Str1)
            
        '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
            'Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing
    
    '//.....................................................................................................
    







Dim Bool1 As Boolean
Dim NameForm_s As String
Dim Msg_s   As String





'EVENTI CLASSE
'______________________________________________________________________________

    
'INIZIALIZZO LA CLASSE

Private Sub Class_Initialize()

    On Error GoTo Err_Class_Initialize

    
            'Reset Variabili Oggetti Form
            m_sxTIPOGGETTO = ""
            m_sxPROPRIETA = ""
            m_sxMETODO = ""
            m_sxEVENTO = ""
            
'USCITA ED ERRORI
'..............................................................
Exit_Class_Initialize:
    Exit Sub

Err_Class_Initialize:
    MsgBox Err.Description
    Resume Exit_Class_Initialize

                                                      
End Sub



'TERMINO LA CLASSE
Private Sub Class_Terminate()

    On Error GoTo Err_Class_Terminate
    
    Dim TITOLO_Get As String


'USCITA ED ERRORI
'..............................................................
Exit_Class_Terminate:
    Exit Sub

Err_Class_Terminate:
    MsgBox Err.Description
    Resume Exit_Class_Terminate


End Sub








'//APRO FILE ATTI DI GESTIONE - CON POWER SCHELL O CON IL DOS
'//###########################################################################################################################//
'//NOTE: apro il file .mdb _
    non ci sono parametri _
    AGGIORNAMENTO DEL = 13.07.2020
'//CODICE----------->APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub

'//AGGIORNAMENTO_DEL_13_07_2020
'//APRI_ATTI_DI_GESTIONE
Public Sub APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub()

'//RECUPERO PARAMETRO DA TABELLA OGGETTI
'//File da ricercare e la path
Dim MyFile_s As String
Dim MyPath_s As String


On Error GoTo Err_APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub

                
  
     
    
    '//APRO SORGENTE FORM FILTRO 66 CAMPI
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire. _
                      Per gestire il prototipo di apertura del file vai al codice ----> @MODELLO@APERTURA@FILE
       
        MyFile_s = "LLPP_ATTI_GESTIONE_Analisi_provisoria.txt"
        '//ATTENZIONE occorre inserire uno spazio finale altrimenti non riconosce il percorso. Se invece è _
           un programma oppure un eseguibile non occorre lo spazio.
        MyPath_s = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\ATTI_TB01_GESTIONE\"
      
       
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub"
        ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
   
      
         
        
        '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
        '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
        '// prima la sessione dos e poi il pdf..
        '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
        Set WshPDF = CreateObject("wscript.shell")
        
        '// Unisco il comando pdf start + path definitiva
        Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
        '//Attivo il comando
        WshPDF.Run Stringa1
        
       
    '//*** fine ***
    '//APRO SORGENTE FORM FILTRO 66 CAMPI
    '//.....................................................................................................
 
                    


    


        
'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub:
    Exit Sub

Err_APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub:
    MsgBox Err.Description
    '//CONTROLLO DELL'ERRORE
    Debug.Print ProceduraAttivaEseguita_s
    Debug.Print ProceduraAttivaEseguita_s

    Resume Exit_APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub

End Sub

'//*** fine ***
'//APRO FILE ATTI DI GESTIONE - CON POWER SCHELL O CON IL DOS
'//###########################################################################################################################//











'//@APRI@FILE@APRI@DIRECTORY_(utilizzando la @CLASSE PER L'APERTURA DEGLI OGGETTI)
'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE CON 3 PARAMETRI-
'//==================================================================================================================//
'//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
'//PARAMETRI            -> passo 3 PARAMETRI 2 stringhe + 1 LONG _
                        par_Path_s As String,           = questa è la path di partenza dove si trova il file _
                        par_NameFile_s As String,       = questo è il nome del file da aprire _
                        par_IDGestione_lng As Long,     = ID di Gestione è il recordo dove è stato attivato l'evento di apertura _
                        par_cmd_b As String             = cmd "DOS" oppure "POWERSHELL" il dos fino a 255 caratteri
                        
'//VALORE_DI_RITORNO    -> Nulla
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.00
'//TODO: manca il 4 parametro che dovrebbe essere l'estensione che per ora non riesco a _
        gestire perche il comando dos non lo permette.
'//ATTENZIONE: se la path + file è troppo lunga di 255 caratteri la procedura non funziona devi ridurre la path _
    OPPURE UTILIZZI LA POWER SHELL DI WINDOS che gestisce oltre i 255 caratteri.



Public Function ApriFilePdf_Interna_s_PFunct(par_Path_s As String, _
                                           par_NameFile_s As String, _
                                           par_IDGestione_lng As Long, _
                                           par_cmd_b As String) As String

    Dim WshShell As Object
    Dim fullFilePath As String
    Dim psCommand As String
    Dim result As Long

    On Error GoTo ApriFilePdf_ConPowerShell_Err



'//CLASSE PER L'APERTURA FILE CON IL DOS O CON IL PAWER SHELL
'//*********************************************************************//
        
          ' Definisco la classe per gestire i file
            '---> Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione
            '---> Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione


        ' Chiamo la funzione della classe passando i parametri
        '---> MyCls_CLASSE_ATTI_Tb01_Gestione.ApriFilePdf_Interna_s_PFunct MyPath_s, MyFile_s, MyParam
        
        
            ' Rilascio l'oggetto e pulisco la memoria
            '---> Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing
            '---> Set fileCollection = Nothing

'//*********************************************************************//



    ' Costruisci il percorso completo del file
    fullFilePath = par_Path_s & par_NameFile_s
    
    '//CONTROLLO PATH
    Debug.Print
    Debug.Print "controllo fullFilePath ---> " & fullFilePath
    
    
            '//UTILIZZO IL DOS   *** ORA SOSPESO DAVA PROBLEMI CON PATH > 255 CARATTERI ***
            '//--------------------------------------------------------------------------------------------------------------------------------//
                  '// CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                  '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                  '// prima la sessione dos e poi il pdf..
                  '// LA STRINGA DI COMANDO DOS DEVE ESERE UNICA: La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                  '//SOSPESO----> Set WshPDF = CreateObject("wscript.shell")
                  
                  '//ATTENZINE LA PATH+FILE > 255 CARATTERI NON FUNZIONE deve rimanere entro le 255 byte altrimenti non funziona.
                  '// CREO UNA STRINGA UNICA, Unisco il comando pdf start + path definitiva altrimenti non funziona.
                  '//SOSPESO----> Stringa1 = "%comspec% /c start " & par_Path_s                               '//Comando di apertura
                  '//Attivo il comando
                  '//SOSPESO----> WshPDF.Run Stringa1
                  
            '//--------------------------------------------------------------------------------------------------------------------------------//

    '//UTILIZZO POWER SCHELL DI WINDOS CHE GESTISCE LE  PATH > 255 CARATTERI SENZA ERRORI
    '//--------------------------------------------------------------------------------------------------------------------------------//
                   
            ' Prepara il comando PowerShell per aprire il file
            psCommand = "powershell -Command Start-Process '" & fullFilePath & "'"
           
            ' Crea un oggetto WScript.Shell
            Set WshShell = CreateObject("WScript.Shell")
        
            ' Esegui il comando PowerShell
            result = WshShell.Run(psCommand, 1, False)
            
            ' Verifica il risultato dell'operazione
            If result <> 0 Then
                MsgBox "Errore nell'apertura del file: " & fullFilePath & vbCrLf & "Codice errore: " & result, vbCritical
            End If
        
            ' Libera la memoria
            Set WshShell = Nothing
            
            
    '//--------------------------------------------------------------------------------------------------------------------------------//

ApriFilePdf_ConPowerShell_Exit:
    Exit Function

ApriFilePdf_ConPowerShell_Err:
    MsgBox "Errore: " & Err.Description, vbCritical
    Resume ApriFilePdf_ConPowerShell_Exit

End Function


'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE CON 3 PARAMETRI- *** fine ***
'//==================================================================================================================//




'//AGGIORNAMENTO_DEL_13_07_2020
'//APRI_ATTI_DI_GESTIONE
Public Sub APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub()

'//RECUPERO PARAMETRO DA TABELLA OGGETTI
'//File da ricercare e la path
Dim MyFile_s As String
Dim MyPath_s As String


On Error GoTo Err_APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub

                
  
     
    
    '//APRO SORGENTE FORM FILTRO 66 CAMPI
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.
       
        MyFile_s = ""
        '//ATTENZIONE occorre inserire uno spazio finale altrimenti non riconosce il percorso. Se invece è _
           un programma oppure un eseguibile non occorre lo spazio.
        MyPath_s = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO\"
      
       
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = ""
        ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
   
      
         
        
        '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
        '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
        '// prima la sessione dos e poi il pdf..
        '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
        Set WshPDF = CreateObject("wscript.shell")
        
        '// Unisco il comando pdf start + path definitiva
        Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
        '//Attivo il comando
        WshPDF.Run Stringa1
        
       
    '//*** fine ***
    '//APRO SORGENTE FORM FILTRO 66 CAMPI
    '//.....................................................................................................
 
                    







        
'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub:
    Exit Sub

Err_APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub:
    MsgBox Err.Description
    Resume Exit_APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub

End Sub

'//*** fine ***
'//APRO LA PATH DEGLI ATTI DI GESTIONE -
'//###########################################################################################################################//


'//AGGIORNAMENTO_DEL_15_08_2020
'//APRO PATH GENERICA = Apre la cartella generica passata come parametro stringa. _
            Routine utilzizabile su tutte le cartelle del progetto, occorre sempre dare la path esplicita _
            e completa di percorso per la sua apertura. Se viene passata la path+file  allora puo essere _
            utilizzata anche per aprime un file generico.
Public Sub APRO_PATH_O_FILE_GENERICI_pSub(par_MyPath_s As String)

'//RECUPERO PARAMETRO DA TABELLA OGGETTI
'//File da ricercare e la path
Dim MyFile_s As String
Dim MyPath_s As String


On Error GoTo Err_APRO_PATH_O_FILE_GENERICI_pSub


    '//APRO PATH O FILE GENERICI
    '//.....................................................................................................
    '//Passo un parametro stringa che può rappresentare una path da aprire esempio : _
       "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO". _
       Se voglio aprire un file  occore passare la PATH+FILE con estensione es. _
       "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO\file.xls"
               
        '//File da ricercare e la path
        'Dim MyFile_s As String
        'Dim MyPath_s As String
        
        'MyFile_s = "LLPP_ATTI_ALLEGATI_ARCHIVIO.RAR"
        'MyPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\LPP_ALLEGATI\"
       
    
         '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
            'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione
    
        '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
           l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
            'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione
        
        '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
            'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
            'MyCls_CLASSE_ATTI_Tb01_Gestione.APRO_PATH_O_FILE_GENERICI_pSub (MyPath_s & MyFile_s)
            
            
        '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
            'Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing
    
    '//.....................................................................................................
                     
            
     '//todo: DA RISISTEMARE L'APERTURA DEI FILE purtroppo un solo parametro la path da errore occorre _
             utilizzare due parametri path + file ma quin non va da errore??. _
             si potrebbe utilizzare la funzione a 3 parametri :  '//chiamo la sub con i parametri =   'CALL (ApriFilePdf_Interna_s_PFunct Path_programma_s, NameFile_s, IDGestione_lng) _
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0) che viene utilizzata in msys
             
      
    
    '//APRO SORGENTE FORM FILTRO 66 CAMPI
    '//.....................................................................................................
    '//Note           : Tramite una Select vengono individuati i valori da restiuire.
       
        MyFile_s = ""
        '//ATTENZIONE occorre inserire uno spazio finale altrimenti non riconosce il percorso. Se invece è _
           un programma oppure un eseguibile non occorre lo spazio.
            'ESEMPIO DI APERTURA DI UNA PATH GENERICA.
            'modello path MyPath_s = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO"
            
        MyFile_s = par_MyPath_s                     'recupero la path passata con il parametro
      
       
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = ""
        ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
   
      
        Debug.Print " MyFile_s = " & MyFile_s
        
        '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
        '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
        '// prima la sessione dos e poi il pdf..
        '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
        Set WshPDF = CreateObject("wscript.shell")
        
        '// Unisco il comando pdf start + path definitiva
        Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura"
        '//Attivo il comando
        WshPDF.Run Stringa1
        
       
    '//*** fine ***
    '//APRO SORGENTE FORM FILTRO 66 CAMPI
    '//.....................................................................................................
 
                    

                      





        
'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_APRO_PATH_O_FILE_GENERICI_pSub:
    Exit Sub

Err_APRO_PATH_O_FILE_GENERICI_pSub:
    MsgBox Err.Description
    Resume Exit_APRO_PATH_O_FILE_GENERICI_pSub

End Sub

'//*** fine ***
'//APRO LA PATH DEGLI ATTI DI GESTIONE -
'//###########################################################################################################################//






'//ITERO ALL.FORMS - OGGETTI FORM DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO FORM PASSATO COME PARAMETRO - _
  EVENTUALMENTE DA UTILIZZARE PER PULIRE LA FORM
'//###########################################################################################################################//
'//NOTE: eseguo il ciclo nell'insieme forms e nell'insieme dei controlli della form filtrata con il nome e condizionata _
dalla if di filtro. Ciclo all'interno delle proprieta delle form e resetto LE CASELLE DI TESTO E LE LABEL.
'//CODICE----------->LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub




Public Sub LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub(par_NameForm_s As String)
    On Error GoTo err_Handler
    Dim frmLoop As Object
    Dim Ctrl As Control
    Dim propLoop As Property
    Dim strFormName As String
    
    '//La variabile BOOLEAN per indicare se esiste
    Dim Bool1 As Boolean
    Dim NameForm_s As String
    Dim Msg_s  As String
    
    
    '//CHIAMO LA ROUTINE DI ITERAZIONE DEI CAMPI TEXT E LABEL DELLA FORM
    '//---------------------------------------------------------------------------------//
    '//CODICE----------->LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub.puliscoTutto
    
        '//ATTIVO LA ROUTINE Call LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub
    
    '//---------------------------------------------------------------------------------//
    
    For Each frmLoop In CurrentProject.AllForms
        
        '//imposto la variabile bool
        Bool1 = False
        '//la form da ricercare nel db
        'NameForm_s = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key"
        
        NameForm_s = par_NameForm_s
        
        '//TROVA QUELLA CORRENTE
        'If frmLoop.NAME = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key" Then
        If frmLoop.Name = par_NameForm_s Then
      
        '//LA FORM CORRENTE
        '//=====================================================================================================//
                
                '//form trovata imposto il nome e  la variabile
                strFormName = frmLoop.Name
                '//imposto la variabile bool a true per individuare che è stata trovata
                Bool1 = True
                
                
                '//APRO LA FORM IN MODALITA DESIGN
                DoCmd.OpenForm strFormName, acDesign, , , , acHidden
                
                '//Qualifico l'oggetto form pe le proprieta
                With forms(strFormName)
                    
                    For Each Ctrl In .Controls
                        
                        DoEvents
                         
                        For Each propLoop In Ctrl.Properties
                        
                             If Ctrl.Name = "TXT_01" Then
                                'Stop
                             End If
                                
                            
                        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                        '//==================================================================================================//
                        '//NOTE : pulisco le caselle di testo impostando la proprieta recordsource a null mentre le label _
                                corrispondenti vengono impostate al nome oggetto label corrente.
                        '//CODICE----------->LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub.puliscoTutto
                        
                        
                            
                            '//PULISCO LA PROPRIETA RECORD SOURCE DELLE CASELLE DI TESTO
                            '//--------------------------------------------------------------------------------------------//
                            '//NELLA CASELLE TEXT riconosciute con il valore ControlSource _
                            viene azzerata la proprieta control source con valore ""
                                If propLoop.Name = "ControlSource" Then
                                    Debug.Print propLoop.Value
                                    'propLoop.value = ""
                                End If
                                
                            '//--------------------------------------------------------------------------------------------//
                            
                                                    
                            
                            '//PULISCO LA PROPRIETA CAPTION DELLE LABEL
                            '//--------------------------------------------------------------------------------------------//
                            '//se è una label riconosciuta con il nome della proprieta _
                            Caption, viene sostituita il valore del contenuto della _
                            etichetta con il nome della label
                                 If propLoop.Name = "Caption" Then
                                            Debug.Print propLoop.Value
                                            '//inserisco il nome label?
                                            'propLoop.value = ctrl.NAME
                                                    
                                                    
                                End If
                                
                            '//--------------------------------------------------------------------------------------------//
                        
                        '//*** FINE ***
                        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                        '//==================================================================================================//
                             
                            '//CONTROLLO oggetto
                            Debug.Print "Form [" & .Name & "] - Control[" & Ctrl.Name & "](" & propLoop.Name & "):" & propLoop.Value
                            
                        Next '//For Each propLoop In ctrl.Properties
                    
                    
                    Next    '//For Each ctrl In .Controls
                
                            '//SALVATAGGIO FORM APERTA IN DESIGN
                            'Richiama la funzione che chiude la maschera aperta
                            DoCmd.SetWarnings False
                            DoCmd.Close acForm, strFormName, acSaveYes
                            DoCmd.SetWarnings True
                            DoCmd.Close acForm, strFormName

                End With '//With Forms(strFormName)
                
                               
        '//=====================================================================================================//
        
                
            
        End If
    Next
    
                '//messaggio di avviso se Bool1=false e cioè non esiste la form
                
                If Bool1 = False Then
                    '//imposto il testo del messaggio
                    Msg_s = "ATTENZIONE LA FORM " & NameForm_s & " NON E' STATA TROVATA IN QUESTO DATABASE!"
                    MsgBox Msg_s, vbExclamation
                End If
    
    
    
 
    Exit Sub
err_Handler:
    Select Case Err.number
        Case 2186
            '//SE LA PROPRIETA NON ESISTE ATTIVO L'ERRORE E RIENTRO NEL CICLO FOR EACH
            'Property not available in design view  SE NON CI SONO LE  PROPRIETA RIPROVA.
            Resume Next
        Case Else
            MsgBox "Err: " & Err.number & " - " & Err.Description
    End Select
End Sub
 
'//*** fine ***
'//ITERO ALL.FORMS - OGGETTI FORM DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO FORM PASSATO COME PARAMETRO - _
  EVENTUALMENTE DA UTILIZZARE PER PULIRE LA FORM
'//###########################################################################################################################//






'//ITERO ALL.MODULES - OGGETTI MODULES DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO PASSATO COME PARAMETRO
'//###########################################################################################################################//
'//NOTE: eseguo il ciclo nell'insieme MODULES e nell'insieme DELLE CLASSI ALLA RICERCA DI QUELLA PASSATA COME PARAMETRO _
   mediante la if di filtro. Ciclo all'interno delle proprieta delle CLASSI E MODULI. - CONTROLLO ESISTENZA MODULI
'//CODICE----------->LoopFormProp_ITERA_OGGETTI_MODULI_E_CLASSI_PSub

'//TODO: DA FINIRE!!!
Public Function LoopModuli_ITERA_OGGETTI_MODULI_E_CLASSI_PFunction(par_NameModulo_s As String) As String
    On Error GoTo err_Handler
    Dim ModuliClassiLoop As Object
    Dim ObjModule As modules
    Dim propLoop As Property
    Dim strModuloName As String
    
    '//La variabile BOOLEAN per indicare se esiste
    Dim Bool1 As Boolean
    Dim NameForm_s As String
    Dim Msg_s  As String
    
    
     
    
     '//CHIAMO LA ROUTINE DI ITERAZIONE DEI MODULI E DELLE CLASSI NEL DB CORRENTE _
       per controlloare se esiste il modulo di classe da attivare
    '//---------------------------------------------------------------------------------//
    '//CODICE----------->LoopFormProp_ITERA_OGGETTI_MODULI_E_CLASSI_PSub.puliscoTutto.IteraTutto
    
       '
       'Dim NWClasse_MODULI_E_CLASSI As CLASSE_ATTI_Tb01_Gestione
               
     
       '//ATTIVO LA ROUTINE
        '
        'Set NWClasse_MODULI_E_CLASSI = New CLASSE_ATTI_Tb01_Gestione
            
            '//imposto il nome della forma da ricercare
            '
         '   Str1 = "MyCls_CLASSE_ATTI_Tb01_Gestione"
            '//ITERO NEGLI OGGETTI DEL DB ALLA RICERCA DEL MODULO O DELLA CLASSE
            '
          '  vV1 = NWClasse_MODULI_E_CLASSI.LoopModuli_ITERA_OGGETTI_MODULI_E_CLASSI_PFunction(Str1)
        
    '//---------------------------------------------------------------------------------//
 
    
    
    For Each ModuliClassiLoop In CurrentProject.AllModules
        
        '//imposto la variabile bool
        Bool1 = False
        
        '//IL MODULO O CLASSE DA ricercare nel db
        strModuloName = par_NameModulo_s
        
      
        
        With ModuliClassiLoop
        '//CONTROLLO oggetto CORRENTE
        Debug.Print "MODULO O CLASSE ITERATO:  [" & ModuliClassiLoop.Name & "]"
        
        End With
                            
        '//TROVA E CONTROLLA CON QUELLO CORRENTE
        If ModuliClassiLoop.Name = par_NameModulo_s Then
      
        '//IL MODULO O CLASSE CORRENTE IN QUESTO DB
        '//=====================================================================================================//
                
                '//form trovata imposto il nome e  la variabile
                strModuloName = ModuliClassiLoop.Name
                '//imposto la variabile bool a true per individuare che è stata trovata
                Bool1 = True
                
                
                '//APRO LA FORM IN MODALITA DESIGN
                'DoCmd.OpenForm strModuloName, acDesign, , , , acHidden
                
                '//Qualifico l'oggetto form pe le proprieta
                With modules(strModuloName)
                   
                    For Each ObjModule In .Parent
                    
                        
                        DoEvents
                         
                        For Each propLoop In ObjModule.Properties
                        
                             If ObjModule.Name = "TXT_01" Then
                                'Stop  TROVATO DA SVILUPPARE!!!!
                             End If
                                
                            
                        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                        '//==================================================================================================//
                        '//NOTE : pulisco le caselle di testo impostando la proprieta recordsource a null mentre le label _
                                corrispondenti vengono impostate al nome oggetto label corrente.
                        '//CODICE----------->LoopModuli_ITERA_OGGETTI_MODULI_E_CLASSI_PFunction.puliscoTutto
                        
                        
                            
                            '//PULISCO LA PROPRIETA RECORD SOURCE DELLE CASELLE DI TESTO
                            '//--------------------------------------------------------------------------------------------//
                            '//NELLA CASELLE TEXT riconosciute con il valore ControlSource _
                            viene azzerata la proprieta control source con valore ""
                                If propLoop.Name = "ControlSource" Then
                                    Debug.Print propLoop.Value
                                    'propLoop.value = ""
                                End If
                                
                            '//--------------------------------------------------------------------------------------------//
                            
                                                    
                            
                            '//PULISCO LA PROPRIETA CAPTION DELLE LABEL - MODIFICARE PER I MODULI E CLASSI
                            '//--------------------------------------------------------------------------------------------//
                            '//se è una label riconosciuta con il nome della proprieta _
                            Caption, viene sostituita il valore del contenuto della _
                            etichetta con il nome della label
                                 If propLoop.Name = "Caption" Then
                                            Debug.Print propLoop.Value
                                            '//inserisco il nome label?
                                            'propLoop.value = ObjModule.NAME
                                                    
                                                    
                                End If
                                
                            '//--------------------------------------------------------------------------------------------//
                        
                        '//*** FINE ***
                        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                        '//==================================================================================================//
                             
                            '//CONTROLLO oggetto
                            Debug.Print "MODULO O CLASSE [" & .Name & "] - Control[" & ObjModule.Name & "](" & propLoop.Name & "):" & propLoop.Value
                            
                        Next '//For Each propLoop In ObjModule.Properties
                    
                    
                    Next    '//For Each ObjModule In .Controls
                
                            '//SALVATAGGIO FORM APERTA IN DESIGN ?? PUO ESSERE UTILIZZATO PER IL MODULO O _
                            PER LA CLASSE !! CONTROLLARE???
                            'Richiama la funzione che chiude la maschera aperta
                            'DoCmd.SetWarnings False
                            'DoCmd.Close acForm, strModuloName, acSaveYes
                            'DoCmd.SetWarnings True
                            'DoCmd.Close acForm, strModuloName

                End With '//With MODULES(strModuloName)
        
        '*** FINE ***
        '//IL MODULO O CLASSE CORRENTE IN QUESTO DB
        '//=====================================================================================================//
        
                
            
        End If
    Next
    
                '//messaggio di avviso se Bool1=false e cioè non esiste la form
                
                If Bool1 = False Then
                    '//imposto il testo del messaggio
                    Msg_s = "ATTENZIONE IL MODULO O LA CLASSE RICERCATA  " & strModuloName & " NON E' STATA TROVATA IN QUESTO DATABASE!"
                    LoopModuli_ITERA_OGGETTI_MODULI_E_CLASSI_PFunction = "OGGETTO NON TROVATO"
                    MsgBox Msg_s, vbExclamation
                End If
    
    
    
 
    Exit Function
err_Handler:
    Select Case Err.number
        Case 2186
            '//SE LA PROPRIETA NON ESISTE ATTIVO L'ERRORE E RIENTRO NEL CICLO FOR EACH
            'Property not available in design view  SE NON CI SONO LE  PROPRIETA RIPROVA.
            Resume Next
        Case Else
            MsgBox "Err: " & Err.number & " - " & Err.Description
    End Select
End Function
 
'//*** fine ***
'//ITERO ALL.MODULES - OGGETTI MODULES DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO PASSATO COME PARAMETRO - _
  EVENTUALMENTE DA UTILIZZARE PER PULIRE I MODULI
'//###########################################################################################################################//
    



'//ITERO IN TUTTI GLI OGGETTI - DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO PASSATO COME PARAMETRO
'//###########################################################################################################################//
'//NOTE: eseguo il ciclo nell'insieme individuato dal paramtro oggetto ALLA RICERCA dell'OGGETTO _
   mediante la if di filtro.  - CONTROLLO ESISTENZA OGGETTO
'//CODICE----------->LoopFormProp_ITERA_OGGETTI_DB_CORRENTE

'****************************************************************************************************************************
'           PROCEDURA DI ITERAZIONE DEGLI OGGETTI DI DATABASE
' i parametri ricevuti sono : par_Scelta db = path db
'
'
'****************************************************************************************************************************

Private Sub ITERA_OGGETTO_DB()




'// FUNZIONE ITERAZIONE OGGETTI, Tabelle, Query, Form, Report, Macro _
    e Moduli
'//====================================================================//
'//MODELLO : di ricerca oggetti, prova la funzione si devono comunque _
            impostare i parametri.
            
            '//Imposto le variabilie e le resetto
            Dim sx_Scelta_db As String
            Dim sx_Name_Tab As String
            Dim sx_Name_Query As String
            Dim sx_Name_FORM As String
            Dim sx_Name_Report As String
            Dim sx_Name_Macro As String
            Dim sx_Name_Modulo As String
            Dim sx_Tipo_Oggetto As String
            Dim ix_Operazione As Integer
            Dim ix_Indice1 As Integer
            Dim bx_Flag_b As Boolean
            
            
            sx_Scelta_db = ""
            sx_Name_Tab = ""
            sx_Name_Query = ""
            sx_Name_FORM = ""
            sx_Name_Report = ""
            sx_Name_Macro = ""
            sx_Name_Modulo = ""
            sx_Tipo_Oggetto = ""
            ix_Operazione = 0
            ix_Indice1 = 0
            bx_Flag_b = False
            
            
            '//chiamo la procedura di iterazione oggetti
            Str1 = SALVATAGGIO_OGGETTI_DB_PFunct(sx_Scelta_db, _
                                      sx_Name_Tab, _
                                      sx_Name_Query, _
                                      sx_Name_FORM, _
                                      sx_Name_Report, _
                                      sx_Name_Macro, _
                                      sx_Name_Modulo, _
                                      sx_Tipo_Oggetto, _
                                      ix_Operazione, _
                                      ix_Indice1, _
                                      bx_Flag_b)
                                      
'//====================================================================//




End Sub



'//###########################################################################################################//
'//####
'//####                             PROCEDURE DI SALVATAGGIO
'//####
'//###########################################################################################################//

'//***********************************************************************************************
'//                 Funzione SALVATAGGIO OGGETTI DATABASE MDB
'//
'//***********************************************************************************************
'//CODICE ---> 2020.02.11_SALVATAGGIO_OGGETTI_DB_PFunct_DB
'PARAMETRI IN IMPUT:
'par_sx_Scelta_db   =database interno od esterno (se la variabile path è valorizzata es. = c:\Gestioni\FILE.mdb _
                     allora viene attivata la II opzione del database esterno);
'par_sx_Name_Tab    =nome tabella da ricercare;
'par_sx_Name_Query  =nome query da ricercare
'par_sTipo_Oggetto  =il tipo oggetto del db sono identificati con la seguente stringa: _
                Cmd_Elenco_Query    = Tutte le query _
                Cmd_Elenco_Tabelle  = Tutte le tabelle _
                Cmd_Elenco_Forms    = Tutte le Form _
                ATTENZIONE mancano ancora i parametri Select per LE FORM, I REPORT, LE MACRO _
                LE PAGINE WEB che sono da integrare.
'tipo di operazione =?  da definire ;
'par_ix_Indice1 =?  da definire ;
'par_bx_Flag_b = True = individuato cosa ??
'
'//OUTPUT : RESTITUISCE una stringa ANCORA DA DEFINRE E IMPLEMENTARE ???
'//MODULO ---> UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB
'//DB MDB ---> UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB

'//MODELLO FUNZIONE ITER OGGETTI
'Public Function SALVATAGGIO_OGGETTI_DB_PFunct(par_sx_Scelta_db As String, _
                                   par_sx_Name_Tab As String, _
                                   par_sx_Name_Query As String, _
                                   par_sx_Name_Form As String, _
                                   par_sTipo_Oggetto As String, _
                                   par_ix_Operazione As Integer, _
                                   par_ix_Indice1 As Integer, _
                                   par_bx_Flag_b As Boolean) As String
'//***********************************************************************************************
'//ATTIVITA ESEGUITE: _
'02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB _
'02.01 - OGGETTI QUERY = CONTROLLO OGGETTO NEL DB: Faccio il Controllo l'esistenza delle TABELLE
'pulendo prima tutta la tabella  oggetti da quelli precedenti.
'03.01 - APRO LA TABELLA DI DATABASE CHE CONTIENE GLI OGGETTI TABELLA
'        Iterando in tutti gli oggetti dell'insieme Tables
'03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB = Itero nel db tra gli oggetti
'       table negli archivi del database e salvo il nome dell'oggetto tabella nel record corrente
'       facendo un ciclo for nell'insieme TableDefs
'03.03 .... definire
'
'//                                     SALVATAGGIO OGGETTI DATABASE
'//{@=============================================================================================================@}
'//CODICE --------------------> SALVATAGGIO_OGGETTI_DB_PFunct.Funzione
'//
Public Function SALVATAGGIO_OGGETTI_DB_PFunct( _
                                   par_sx_Scelta_db As String, _
                                   par_sx_Name_Tab As String, _
                                   par_sx_Name_Query As String, _
                                   par_sx_Name_FORM As String, _
                                   par_sx_Name_Report As String, _
                                   par_sx_Name_MACRO As String, _
                                   par_sx_Name_Modulo As String, _
                                   par_sTipo_Oggetto As String, _
                                   par_ix_Operazione As Integer, _
                                   par_ix_Indice1 As Integer, _
                                   par_bx_Flag_b As Boolean) As String


        On Error GoTo Err_SALVATAGGIO_OGGETTI_DB_PFunct

        '// FUNZIONE ITERAZIONE OGGETTI, Tabelle, Query, Form, Report, Macro _
            e Moduli
        '//====================================================================//
        'Str1 = SALVATAGGIO_OGGETTI_DB_PFunct(sScelta_db, _
                                    sName_Tab, _
                                    sName_Query, _
                                    sName_Form, _
                                    sName_Report, _
                                    sName_Macro, _
                                    sName_Modulo, _
                                    sTipo_Oggetto, _
                                    iOperazione, _
                                    iIndice1, _
                                    Flag_b)
        '//====================================================================//


        '------------------------------------------------------------------------------------------
            '1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB
         
                    'OPEN DB: Apre il database scelto per la creazione o cancellazione
                    '         con il parametro Option = false per indicare che l'apertura è condivisa /
                    '                          Option = True apertura esclusiva
                    
                If par_sx_Scelta_db = "" Then
                        'apro il db corrente se non viene indicato un database esterno
                        Set Dbs = CurrentDb
                    Else
                        'apro un database esterno con la path >0 passata con i parametri
                        Set Dbs = OpenDatabase(par_sx_Scelta_db, False)
                End If
            
            Select Case par_sTipo_Oggetto
                
                '//- CONTROLLO TABELLE -
                '//CHIAVE =====> modulo:controllo-tabelle
                Case Is = "Cmd_Elenco_Tabelle"
                        
                        '02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB
                        '______________________________________________________________________________
                        '02.01 - OGGETTI QUERY
                        
                            
                            'pulisco la tabella degli OGGETTI TABELLA
                            sSql = ""
                            sSql = sSql & "DELETE MSsys_TABELLE_Qry01_01_TUTTE.*"
                            sSql = sSql & "FROM MSsys_TABELLE_Qry01_01_TUTTE "
                            sSql = sSql & "WITH OWNERACCESS OPTION;"

                            CurrentDb.Execute (sSql)
                            
                            '03.01 - APRO LA TABELLA
                            'apro la tabella per il salvataggio delle query + effettuo il _
                            controllo sulla stringa sql
                            Debug.Print "---------------------------------------------------------------"
                            Debug.Print "CONTROLLO STRINGA SQL                                          "
                            Debug.Print
                            Set ApriRs = Dbs.OpenRecordset("MSsys_TABELLE_Qry01_01_TUTTE")
                            
                        
                            With Dbs
                                
                                '//reset
                                iCount = 0
                                
                                '//stampa di controllo
                                Debug.Print .TableDefs.count & _
                                    " TableDefs in database  " & sx_Scelta_db
                                
                                
                                For intCiclo = 0 To .TableDefs.count - 1
                                        Debug.Print "  " & .TableDefs(intCiclo).Name
                                   
                                        '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                        '................................................................
                                            ApriRs.AddNew
                                            ApriRs.Fields("Name1_s") = .TableDefs(intCiclo).Name
                                            
                                            '//NRO di oggetto
                                            iCount = iCount + 1         '//aggiorno il contatore oggetti delle tabelle
                                            ApriRs.Fields("NRO_OGGETTO_i") = iCount
                                            '//Tipo di oggetto caricato = FORM
                                            ApriRs.Fields("TIPOGGETTO_s") = "TABLE"
                                            
                                            '//SALVATAGGIO DEGLI ATTRIBUTI DELLA TABELLA - proprieta Attributes
                                            ApriRs.Fields("Attribute_lng") = .TableDefs(intCiclo).Attributes
                                            
                                                '//CONTROLLO OGGETTO TABELLA
                                                '//-----------------------------------------------------------
                                                    
                                                    '//controllo degli attributi
                                                    Debug.Print "LA TABELLA "
                                                    Debug.Print .TableDefs(intCiclo).Name
                                                    
                                                    Debug.Print "controllo attributi della tabella"
                                                    Debug.Print .TableDefs(intCiclo).Attributes
                                                    
                                                '//-----------------------------------------------------------
                                            ApriRs.Update
                                        '................................................................
                                    
                                

                                    '03.03 .... definire
                                    If .TableDefs(intCiclo).Name = par_sx_Name_Tab Then
                                        Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                        SALVATAGGIO_OGGETTI_DB_PFunct = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                        GoTo Exit_SALVATAGGIO_OGGETTI_DB_PFunct  'esci
                                    Else
                                        Bln1 = False   ' bx_Flag_b False = non trovato TABELLA
                                        SALVATAGGIO_OGGETTI_DB_PFunct = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                    End If
                                    
                                Next intCiclo
        
                            End With
                           
               ' - CONTROLLO QUERY -
                Case Is = "Cmd_Elenco_Query"
                    '//CHIAVE =====> modulo:controllo-Query
                    
                        '//ITERAZIONE QUERYDEFS <<LE QUERY>>
                        '//====================================================================================================//
                        'OGGETTI QUERY
                        'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della QUERY
                        
                        
                        'pulisco la tabella delle query prima del salvataggio dei nuovi nomi
                        CurrentDb.Execute ("DELETE MSsys_QUERY_Qry01_01_TUTTE.* FROM MSsys_QUERY_Qry01_01_TUTTE;")
                        
                        'apro la tabella per il salvataggio delle query
                        Set ApriRs = Dbs.OpenRecordset("MSsys_QUERY_Qry01_01_TUTTE") '// perche non funziona??? da errore
                        
                        'itero nel db
                                With Dbs
                                    
                                    '//RESET VARIABILI CICLO
                                    iCount = 0      '//contatore numero oggetti query
                                    
                                    
                                    Debug.Print .QueryDefs.count & _
                                        " QueryDefs in database  " & sx_Scelta_db
                                        
                                        
                                '//CICLO OGGETTI QUERY
                                '//------------------------------------------------------------------------------//
                                For intCiclo = 0 To .QueryDefs.count - 1
                                        
                                        'controlli
                                        '................................................................
                                            
                                
                                                Debug.Print "-----------------------------------------------------------"
                                                Debug.Print
                                                            '//stampo il nome della query + il contentuto sql
                                                            Debug.Print " stampo il NOME DELLA QUERY"
                                                            Debug.Print "  " & .QueryDefs(intCiclo).Name
                                                           
                                                            '//AGGIORNAMENTO_DEL_28_02_2020
                                                            'Stampo il contenuto sql della query
                                                            Debug.Print
                                                            Debug.Print "  " & .QueryDefs(intCiclo).SQL
                                                            
                                                            '//AGGIORNAMENTO_DEL_28_02_2020
                                                            '//stampo il tipo di query scelta direttamente
                                                            Debug.Print "IL TIPO DELLA Query"
                                                            Debug.Print "  " & .QueryDefs(intCiclo).Type
                                                
                                                Debug.Print "-----------------------------------------------------------"
                                                Debug.Print
                                                
                                            
                                        '................................................................
                                            
                                            'SALVATAGGIO NOME QUERY
                                            '................................................................
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//Aggiorno il numero oggetto query salvati nel campo della tabella _
                                                dei salvataggi del nro query.
                                                iCount = iCount + 1         '//aggiorno il contatore oggetti delle query
                                              
                                                
                                                ApriRs.AddNew
                                                '//aggiugno alla tabella il nome della query + il contenuto sql
                                                ApriRs.Fields("Name1_s") = .QueryDefs(intCiclo).Name
                                                ApriRs.Fields("IdOrd_lng") = iCount
                                                
                                                '//Tipo di oggetto caricato = QUERY
                                                ApriRs.Fields("TIPOGGETTO_s") = "QUERY"
                                                
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//SQL + TIPO DI QUERY
                                                ApriRs.Fields("Expression_m") = .QueryDefs(intCiclo).SQL
                                                            
                                                
                                                ApriRs.Update
                                                
                                                
                                                
                                                
                                                '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                        '................................................................
                                               
                                                
                                                
                                            '................................................................
                                            
                                        'controllo nome query passata con parametro
                                        If .QueryDefs(intCiclo).Name = par_sx_Name_Query Then
                                            
                                            Bln1 = True    ' bx_Flag_b True = trovato QUERY
                                            
                                            'stampo il contenuto della query scelta direttamente _
                                            solo se esiste la query ricercata.
                                            Debug.Print "  " & .QueryDefs(par_sx_Name_Query).SQL
                                            
                                            
                                            'recupero il contenuto della query
                                            SALVATAGGIO_OGGETTI_DB_PFunct = .QueryDefs(intCiclo).SQL
                                            
                                            
                                            GoTo Exit_SALVATAGGIO_OGGETTI_DB_PFunct  'esci
                                        Else
                                            Bln1 = False   ' bx_Flag_b False = non trovato QUERY
                                            SALVATAGGIO_OGGETTI_DB_PFunct = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                        End If
                                        
                                    Next intCiclo
                                    
                                '//CICLO OGGETTI QUERY  --> Next intCiclo *** fine ***
                                '//------------------------------------------------------------------------------//
                                    
                                End With
                        '//ITERAZIONE QUERYDEFS <<LE QUERY>>   *** FINE ***
                        '//====================================================================================================//
                        

            
               'CICLO NEGLI OGGETTI FORMS
               Case Is = "Cmd_Elenco_Forms"
               '//CHIAVE =====> modulo:controllo-forms
                
               
                 
   
                        '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                        '//==============================================================================================================//
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS
                        
                            
                            'Dim obj As AccessObject, obj_Oggetto As Object
                            Set obj_Oggetto = Application.CurrentProject
   
                             '//RESET DELLE VARIABILI
                            intCiclo = 0
                            iCount = 0
                           
                            'NOME form prorvvisoria ca controllare, poi cancellare
                            par_sx_Name_FORM = "form ??? da definire"

                            
                            'apro la tabella per il salvataggio delle query
                            Set ApriRs = Dbs.OpenRecordset("MSsys_FORM_Qry01_01_TUTTE")
                            
                            
                            '//RESET DELLE VARIABILI
                            intCiclo = 0
                            iCount = 0
                            
                            'pulisco la tabella delle FORMS prima del salvataggio dei nuovi nomi
                            CurrentDb.Execute ("DELETE MSsys_FORM_Qry01_01_TUTTE.* FROM MSsys_FORM_Qry01_01_TUTTE;")
                            
                            
                                                        
                        '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                        '______________________________________________________________________________
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS
                            
                                                       
                            
                        'itero nel db
                        With Dbs
                                    
                                    
                            
                        '//ITERAZIONE Forms DEFS <<LE FORMS>> - CONTROLLO FORM -
                        '//______________________________________________________________________________
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della Forms
                                
                                For Each obj In obj_Oggetto.AllForms
                                        
                                        
                                        'SALVATAGGIO NOME Forms
                                        '................................................................
                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                        '//Aggiorno il numero oggetto Forms  salvati nel campo della tabella _
                                        dei salvataggi del nro Forms.
                                          
                                          ApriRs.AddNew
                                            
                                            'CONTROLLO DEI NOMI DELE FORMS
                                            '................................................................
                                                        
                                                    '//Search for open AccessObject objects in AllForms collection. _
                                                    con la proprieta IsLoad = esegui il controllo delle Forms Aperti _
                                                    mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                    Report del Db. Per le Forms APERTI O CARICATI  viene impostata la _
                                                    proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                    FALSE. _
                                                    È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                    un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.
                                                                
                                                                
                                                                Debug.Print "  " & obj.Name
                                                                Debug.Print "  " & obj.Type
                                                                
                                                                 '//Tipo di oggetto caricato = QUERY
                                                                 ApriRs.Fields("TIPOGGETTO_s") = "FORMS"
                                                
                                               
                                                    
                                                    '//obj.Type = 2 LA FORMS
                                                    If obj.Type = 2 Then
    
                                                            If obj.IsLoaded = True Then
                                                    
                                                                   '//aggiugno alla tabella il nome della FORM solo se oggetto di tipo 2
                                                                    ApriRs.Fields("Name1_s") = obj.Name
                                            
                                                                    Debug.Print "FORM APERTA = " & obj.Name
                                                                    Debug.Print "  " & obj.Name
                                                                    Debug.Print "Nro : " & iCount
                                                                    ApriRs.Fields("PROPRIETA_s") = "True"
                                                                    ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle Forms
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
        
                                                                    
                                                            Else
                                                                    
                                                                    '//aggiugno alla tabella il nome della FORM solo se oggetto di tipo 2
                                                                    ApriRs.Fields("Name1_s") = obj.Name
                                            
                                                                    
                                                                    Debug.Print "FORM NON APERTO ---> " & obj.Name
                                                                    ApriRs.Fields("PROPRIETA_s") = "FALSE"
                                                                    ApriRs.Fields("VALOREPROPRIETA_m") = "FORM CHIUSA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle Forms
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
                                                            End If
                                                    End If
                                                              
                                                                          
                                            '................................................................
                                            
                                          ApriRs.Update
                                        '................................................................
                                      
                                Next obj
                                
                            End With

            '//ITERAZIONE FORMS DEFS <<LE FORMS>>    *** FINE ***
            '//==============================================================================================================//

                            '**************************************************************
                            
                            
                                                        
              'CICLO NEGLI OGGETTI REPORTS
               Case Is = "Cmd_Elenco_Reports"
                '//CHIAVE =====> modulo:controllo-Report
                '//Aggiornamento 22 MARZO 2020
                
                        '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                        '//______________________________________________________________________________
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS
                            
                            
                            
                            'Dim obj As AccessObject, obj_Oggetto As Object
                            Set obj_Oggetto = Application.CurrentProject
   
                            '//RESET DELLE VARIABILI
                            intCiclo = 0
                            iCount = 0
                            
                            'pulisco la tabella delle FORMS prima del salvataggio dei nuovi nomi
                            CurrentDb.Execute ("DELETE MSsys_REPORT_Qry01_01_TUTTI.* FROM MSsys_REPORT_Qry01_01_TUTTI;")
                            
                            'apro la tabella per il salvataggio dei REPORT
                            Set ApriRs = Dbs.OpenRecordset("MSsys_REPORT_Qry01_01_TUTTI")
                            
                                                        
                        '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                        '______________________________________________________________________________
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS
                            
                                                       
                            
                        'itero nel db
                        With Dbs
                                    
                                    
                            
                        '//ITERAZIONE REPORTS DEFS <<I REPORT>> - CONTROLLO REPORT -
                        '//______________________________________________________________________________
                        '//OGGETTI REPORTS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della REPORTS
                                
                                For Each obj In obj_Oggetto.AllReports
                                        
                                        
                                        'SALVATAGGIO NOME REPORTS
                                        '................................................................
                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                        '//Aggiorno il numero oggetto REPORTS  salvati nel campo della tabella _
                                        dei salvataggi del nro REPORTS.
                                       
                                    
                                          
                                          ApriRs.AddNew
                                            '//aggiugno alla tabella il nome della query + il contenuto sql
                                            ApriRs.Fields("Name1_s") = obj.Name
                                            
                                            
                                            
                                            
                                            'CONTROLLO DEI NOMI DEI REPORT
                                            '................................................................
                                                        
                                                    '//Search for open AccessObject objects in AllReports collection. _
                                                    con la proprieta IsLoad = esegui il controllo dei Reports Aperti _
                                                    mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                    Report del Db. Per i reports APERTI O CARICATI  viene impostata la _
                                                    proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                    FALSE. _
                                                    È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                    un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.
                                                                
                                                                
                                                                Debug.Print "  " & obj.Name
                                                                Debug.Print "  " & obj.Type
                                                                
                                                                '//Tipo di oggetto caricato = REPORTS
                                                                ApriRs.Fields("TIPOGGETTO_s") = "REPORTS"
                                                    
                                                    '//obj.Type = 3 LA FORMS
                                                    If obj.Type = 3 Then
    
                                                            If obj.IsLoaded = True Then
                                                    
                                                                   
                                                                    Debug.Print "REPORT APERTO = " & obj.Name
                                                                    Debug.Print "  " & obj.Name
                                                                    Debug.Print "Nro : " & iCount
                                                                    ApriRs.Fields("PROPRIETA_s") = "True"
                                                                    ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle REPORTS
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
        
                                                                    
                                                            Else
                                                                    Debug.Print "REPORT NON APERTO ---> " & obj.Name
                                                                    ApriRs.Fields("PROPRIETA_s") = "FALSE"
                                                                    ApriRs.Fields("VALOREPROPRIETA_m") = "FORM CHIUSA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle REPORTS
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
                                                            End If
                                                    End If
                                                               
                                            '................................................................
                                            
                                            
                                            
                                            
                                          
                                          ApriRs.Update
                                        '................................................................
                                        
                                        
                                        
                                    
                                Next obj
                                
                            End With

   
                 
                        '//ITERAZIONE REPORTS DEFS <<I REPORT>>
                        '//______________________________________________________________________________
                            
                            
               'CICLO NEGLI OGGETTI MACRO - CONTROLLO MACRO -
               Case Is = "Cmd_Elenco_MACRO"
                '//CHIAVE =====> modulo:controllo-Macro
                '//Aggiornamento 22 MARZO 2020
                
                        '//ITERAZIONE MACRO DEFS <<LE MACRO>>
                        '//______________________________________________________________________________
                        '//OGGETTI MACRO
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO
                            
                            
                            
                            'Dim obj As AccessObject, obj_Oggetto As Object
                            Set obj_Oggetto = Application.CurrentProject
   
                            '//RESET DELLE VARIABILI
                            intCiclo = 0
                            iCount = 0
                            
                            'pulisco la tabella delLE MACRO prima del salvataggio dei nuovi nomi
                            CurrentDb.Execute ("DELETE MSsys_MACRO_Qry01_01_TUTTE.* FROM MSsys_MACRO_Qry01_01_TUTTE;")
                            
                            'apro la tabella per il salvataggio dei Macro
                            Set ApriRs = Dbs.OpenRecordset("MSsys_MACRO_Qry01_01_TUTTE")
                            
                                                        
                        '//ITERAZIONE MACRO DEFS <<LE MACRO>>
                        '______________________________________________________________________________
                        '//OGGETTI MACRO
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO
                            
                                                       
                            
                        'itero nel db
                        With Dbs
                                    
                                    
                            
                        '//ITERAZIONE MACRO DEFS <<I Macro>>
                        '//______________________________________________________________________________
                        '//OGGETTI MACRO
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO
                                
                                For Each obj In obj_Oggetto.AllMacros
                                        
                                        
                                        'SALVATAGGIO NOME MACRO
                                        '................................................................
                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                        '//Aggiorno il numero oggetto MACRO  salvati nel campo della tabella _
                                        dei salvataggi del nro MACRO.
                                       
                                    
                                          
                                          ApriRs.AddNew
                                            '//aggiugno alla tabella il nome della query + il contenuto sql
                                            ApriRs.Fields("Name1_s") = obj.Name
                                            
                                                    
                                            
                                            
                                            
                                            'CONTROLLO DEI NOMI DEI Macro
                                            '................................................................
                                                        
                                                    '//Search for open AccessObject objects in AllMACRO collection. _
                                                    con la proprieta IsLoad = esegui il controllo dei MACRO Aperti _
                                                    mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                    Macro del Db. Per i MACRO APERTI O CARICATI  viene impostata la _
                                                    proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                    FALSE. _
                                                    È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                    un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.
                                                                
                                                                
                                                                Debug.Print "  " & obj.Name
                                                                Debug.Print "  " & obj.Type
                                                                
                                                                '//Tipo di oggetto caricato = MACRO
                                                                ApriRs.Fields("TIPOGGETTO_s") = "MACRO"
                                                
                                                    
                                                    '//obj.Type = 4 LA MACRO
                                                    If obj.Type = 4 Then
    
                                                           
                                                    
                                                                   
                                                                    Debug.Print "Macro APERTO = " & obj.Name
                                                                    Debug.Print "  " & obj.Name
                                                                    Debug.Print "Nro : " & iCount
                                                                    'ApriRs.Fields("PROPRIETA_s") = "True"
                                                                    'ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle MACRO
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
        
                                                                    
                                                        
                                                    End If
                                                               
                                            '................................................................
                                            
                                            
                                            
                                            
                                          
                                          ApriRs.Update
                                        '................................................................
                                        
                                        
                                        
                                    
                                Next obj
                                
                            End With

   
                 
                        '//ITERAZIONE MACRO DEFS <<I Macro>>
                        '//______________________________________________________________________________
         
                            
                    
                    Case Else
                            MsgBox "Scelta comando non effettuata", vbExclamation
                             GoTo Exit_SALVATAGGIO_OGGETTI_DB_PFunct
                End Select
                            'chiudo i rs aperti + Rilascio la memoria
                            ApriRs.Close
                            Set ApriRs = Nothing
                            
                            
                                    
'------------------------------------------------------------------------------
'                       FINE FUNCTIONE E GESTIONE ERRORI

Exit_SALVATAGGIO_OGGETTI_DB_PFunct:
    Dbs.Close
    Set Dbs = Nothing
    Exit Function


Err_SALVATAGGIO_OGGETTI_DB_PFunct:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.number & " - " & Err.Description, vbCritical, "SALVATAGGIO_OGGETTI_DB_PFunct"
    Resume Exit_SALVATAGGIO_OGGETTI_DB_PFunct
    
End Function

'//*** FINE ***
'//                                     SALVATAGGIO OGGETTI DATABASE
'//{@=============================================================================================================@}


 

'//*** fine ***
'//ITERO IN TUTTI GLI OGGETTI - DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO PASSATO COME PARAMETRO
'//###########################################################################################################################//




'****************************************************************************************************************************
'           PROCEDURA DI ITERAZIONE DEGLI OGGETTI DI DATABASE PER IL CONTROLLO ESISTENZA OGGETTI
' i parametri ricevuti sono : par_Scelta db = path db
'
'
'****************************************************************************************************************************



'//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Function.Salvataggio_ATTIVITA
'//SALVATAGGIO PARAMETRI NELLA TABELLA TMP
'//###########################################################################################################################//
'//NOTE: Nella tabella <<GEST_Tb01_N01_PARAM_OGGETTI_TMP>> _
viene effettuato il salvataggio dei parametri. I parametri in entrata sono i seguenti :
'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
'par_sxTIPOGGETTO       =
'par_sxATTIVITAEVENTO   =
'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
Public Sub SALVO_ATTIVITA_CMB_02_IN_TMP_Sub(par_sxCODOGGETTO As String, _
                                         par_sxTIPOGGETTO As String, _
                                         par_sxATTIVITAEVENTO As String, _
                                         par_sxCODATTIVITA As String)




 
    '//CHIAMO LA ROUTINE SALVO_ATTIVITA_CMB_02_IN_TMP_Sub DELLA CLASSE OGGETTI FORM
    '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Salvataggio_ATTIVITA.Attivazione
    '//---------------------------------------------------------------------------------//
    '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_FORM
    '//PARAMETRI            : sono 4 di seguito indicati
            'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
            'par_sxTIPOGGETTO       =
            'par_sxATTIVITAEVENTO   =
            'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
    '//NOTA  :Salvo i parametri del codice oggetto e del codice attiva nella tabella tmp solo nel caso in cui _
            sia IL CODICE OGGETTO che il CODICE ATTIVITA sono VALORIZZATI altrimenti la classe ricevente emette _
            un msg di avviso.
        
    
    
                'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                'Assegno e creo l'oggetto di classe
                'Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As Classe_GESTIONE_OGGETTI_FORM
                'Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New Classe_GESTIONE_OGGETTI_FORM
                
                
                'PARAMETRI DI ESTRAZIONE
                'Attivo il metodo della classe rappresentata dalla routine passando
                'i 4 parametri necessari
                'sxCODOGGETTO = s
                'sxTIPOGGETTO = s
                'sxATTIVITAEVENTO = s
                'sxCODATTIVITA = s
                'MyCls_CLASSE_GESTIONE_OGGETTI_FORM.SALVO_ATTIVITA_CMB_02_IN_TMP_Sub sxCODOGGETTO, sxTIPOGGETTO, sxATTIVITAEVENTO, sxCODATTIVITA
                
                'Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing
                                            
        '//---------------------------------------------------------------------------------//
    
    



On Error GoTo Err_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub



                '//CONTROLLO PARAMETRI RICEVENTI
                '//---------------------------------------------------------------------------------//
                '//NOTA : se almeno i due parametri più importanti, ossia CODICE ATTIVITA E CODICE
                '//     OGGETTO non sono valorizzati, viene emesso un messaggio di avviso e di uscita.
                    If sxCODOGGETTO > "" And sxCODATTIVITA > "" Then
                        MsgBox "ATTENZIONE I PARAMETRI CODICE OGGETTO E CODICE ATTIVITA NON SONO VALORIZZATI", vbExclamation, "CLASSE - CONTROLLO PARAMETRI"
                    End If
                '//---------------------------------------------------------------------------------//
                    


                    
                    'SALVATAGGIO ATTIVITA NELLA TABELLA GEST_Tb01_N01_PARAM_OGGETTI_TMP
                    '___________________________________________________________________
                        'Stringa ssql
                        'di salvataggio nella tabella tmp della attivita da svolgere
                        sSql = ""
                        sSql = sSql & "UPDATE GEST_Tb01_N01_PARAM_OGGETTI_TMP "
                        sSql = sSql & "SET GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODATTIVITA = '" & sxCODATTIVITA & "', "
                                                'sSql = sSql & "SET GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODATTIVITA = '" & sxCODATTIVITA & "'; "
                        sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODOGGETTO = '" & par_sxCODOGGETTO & "'; "

                        
                        'controlo
                        Debug.Print sSql
                        
                        'comando eseguito
                        CurrentDb.Execute (sSql)
                    '___________________________________________________________________
                    

        
'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub:
    Exit Sub

Err_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub:
    MsgBox Err.Description
    Resume Exit_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub

End Sub

'//*** FINE ***
'//SALVATAGGIO PARAMETRI NELLA TABELLA TMP
'//###########################################################################################################################//




'//###########################################################################################################//
'//####
'//####                             PROCEDURE DI SALVATAGGIO *** FINE ***
'//####
'//###########################################################################################################//









'//###########################################################################################################//
'//####
'//####                             PROCEDURE DI RICERCA
'//####
'//###########################################################################################################//

'//***********************************************************************************************
'//                 Funzione RICERCA OGGETTI DATABASE MDB
'//
'//***********************************************************************************************
'//CODICE ---> 2020.02.11_RICERCA_OGGETTI_DB_PFunct_DB
'PARAMETRI IN IMPUT:
'par_sx_Scelta_db   =database interno od esterno (se la variabile path è valorizzata es. = c:\Gestioni\FILE.mdb _
                     allora viene attivata la II opzione del database esterno);
'par_sx_Name_Tab    =nome tabella da ricercare;
'par_sx_Name_Query  =nome query da ricercare
'par_sTipo_Oggetto  =il tipo oggetto del db sono identificati con la seguente stringa: _
                Cmd_Elenco_Query    = Tutte le query _
                Cmd_Elenco_Tabelle  = Tutte le tabelle _
                Cmd_Elenco_Forms    = Tutte le Form _
                ATTENZIONE mancano ancora i parametri Select per LE FORM, I REPORT, LE MACRO _
                LE PAGINE WEB che sono da integrare.
'tipo di operazione =?  da definire ;
'par_ix_Indice1 =?  da definire ;
'par_bx_Flag_b = True = individuato cosa ??
'
'//OUTPUT : RESTITUISCE UN MESSAGGIO "OGGETTO TROVATO NEL DB" O "NON TROVATO NEL DB"
'//MODULO ---> UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB
'//DB MDB ---> UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB

'//***********************************************************************************************
'//ATTIVITA ESEGUITE: _
'02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB _
'02.01 - OGGETTI QUERY = CONTROLLO OGGETTO NEL DB: Faccio il Controllo l'esistenza delle TABELLE
'pulendo prima tutta la tabella  oggetti da quelli precedenti.
'03.01 - APRO LA TABELLA DI DATABASE CHE CONTIENE GLI OGGETTI TABELLA
'        Iterando in tutti gli oggetti dell'insieme Tables
'03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB = Itero nel db tra gli oggetti
'       table negli archivi del database e salvo il nome dell'oggetto tabella nel record corrente
'       facendo un ciclo for nell'insieme TableDefs
'03.03 .... definire
'
'//                                     SALVATAGGIO OGGETTI DATABASE
'//{@=============================================================================================================@}
'//CODICE --------------------> RICERCA_OGGETTI_DB_PFunct.Funzione
'//
Public Function RICERCA_OGGETTI_DB_PFunct( _
                                   par_sx_Scelta_db As String, _
                                   par_sx_Name_Tab As String, _
                                   par_sx_Name_Query As String, _
                                   par_sx_Name_FORM As String, _
                                   par_sx_Name_Report As String, _
                                   par_sx_Name_MACRO As String, _
                                   par_sx_Name_Modulo As String, _
                                   par_sTipo_Oggetto As String, _
                                   par_ix_Operazione As Integer, _
                                   par_ix_Indice1 As Integer, _
                                   par_bx_Flag_b As Boolean) As String


        On Error GoTo Err_RICERCA_OGGETTI_DB_PFunct

       
        '------------------------------------------------------------------------------------------
            '1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB
         
                    'OPEN DB: Apre il database scelto per la creazione o cancellazione
                    '         con il parametro Option = false per indicare che l'apertura è condivisa /
                    '                          Option = True apertura esclusiva
                    
                If par_sx_Scelta_db = "" Then
                        'apro il db corrente se non viene indicato un database esterno
                        Set Dbs = CurrentDb
                    Else
                        'apro un database esterno con la path >0 passata con i parametri
                        Set Dbs = OpenDatabase(par_sx_Scelta_db, False)
                End If
            
            Select Case par_sTipo_Oggetto
                
                '//- CONTROLLO TABELLE -
                '//CHIAVE =====> modulo:controllo-tabelle
                Case Is = "Cmd_Elenco_Tabelle"
                        
                        '02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB
                        '______________________________________________________________________________
                        '02.01 - OGGETTI QUERY
                        
                            
                            'pulisco la tabella degli OGGETTI TABELLA
                            sSql = ""
                            sSql = sSql & "DELETE MSsys_TABELLE_Qry01_01_TUTTE.*"
                            sSql = sSql & "FROM MSsys_TABELLE_Qry01_01_TUTTE "
                            sSql = sSql & "WITH OWNERACCESS OPTION;"

                            CurrentDb.Execute (sSql)
                            
                            '03.01 - APRO LA TABELLA
                            'apro la tabella per il salvataggio delle query + effettuo il _
                            controllo sulla stringa sql
                            Debug.Print "---------------------------------------------------------------"
                            Debug.Print "CONTROLLO STRINGA SQL                                          "
                            Debug.Print
                            Set ApriRs = Dbs.OpenRecordset("MSsys_TABELLE_Qry01_01_TUTTE")
                            
                        
                            With Dbs
                                
                                '//reset
                                iCount = 0
                                
                                '//stampa di controllo
                                Debug.Print .TableDefs.count & _
                                    " TableDefs in database  " & sx_Scelta_db
                                
                                
                                For intCiclo = 0 To .TableDefs.count - 1
                                        Debug.Print "  " & .TableDefs(intCiclo).Name
                                   
                                        '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                        '................................................................
                                            ApriRs.AddNew
                                            ApriRs.Fields("Name1_s") = .TableDefs(intCiclo).Name
                                            
                                            '//NRO di oggetto
                                            iCount = iCount + 1         '//aggiorno il contatore oggetti delle tabelle
                                            ApriRs.Fields("NRO_OGGETTO_i") = iCount
                                            '//Tipo di oggetto caricato = FORM
                                            ApriRs.Fields("TIPOGGETTO_s") = "TABLE"
                                            
                                            '//SALVATAGGIO DEGLI ATTRIBUTI DELLA TABELLA - proprieta Attributes
                                            ApriRs.Fields("Attribute_lng") = .TableDefs(intCiclo).Attributes
                                            
                                                '//CONTROLLO OGGETTO TABELLA
                                                '//-----------------------------------------------------------
                                                    
                                                    '//controllo degli attributi
                                                    Debug.Print "LA TABELLA "
                                                    Debug.Print .TableDefs(intCiclo).Name
                                                    
                                                    Debug.Print "controllo attributi della tabella"
                                                    Debug.Print .TableDefs(intCiclo).Attributes
                                                    
                                                '//-----------------------------------------------------------
                                            ApriRs.Update
                                        '................................................................
                                    
                                

                                    '03.03 .... definire
                                    If .TableDefs(intCiclo).Name = par_sx_Name_Tab Then
                                        Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                        RICERCA_OGGETTI_DB_PFunct = "QUERY_TROVATA_NEL_DB" 'riassegno alla funzione il valore del bx_Flag_b
                                        GoTo Exit_RICERCA_OGGETTI_DB_PFunct  'esci
                                    Else
                                        Bln1 = False   ' bx_Flag_b False = non trovato TABELLA
                                        RICERCA_OGGETTI_DB_PFunct = "QUERY_NON_TROVATA_NEL_DB" 'riassegno alla funzione il valore del bx_Flag_b
                                    End If
                                    
                                Next intCiclo
        
                            End With
                           
               ' - CONTROLLO QUERY -
                Case Is = "Cmd_Elenco_Query"
                    '//CHIAVE =====> modulo:controllo-Query
                    
                        '//ITERAZIONE QUERYDEFS <<LE QUERY>>
                        '//====================================================================================================//
                        'OGGETTI QUERY
                        'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della QUERY
                        
                        
                        'pulisco la tabella delle query prima del salvataggio dei nuovi nomi
                        CurrentDb.Execute ("DELETE MSsys_QUERY_Qry01_01_TUTTE.* FROM MSsys_QUERY_Qry01_01_TUTTE;")
                        
                        'apro la tabella per il salvataggio delle query
                        Set ApriRs = Dbs.OpenRecordset("MSsys_QUERY_Qry01_01_TUTTE") '// perche non funziona??? da errore
                        
                        'itero nel db
                                With Dbs
                                    
                                    '//RESET VARIABILI CICLO
                                    iCount = 0      '//contatore numero oggetti query
                                    
                                    
                                    Debug.Print .QueryDefs.count & _
                                        " QueryDefs in database  " & sx_Scelta_db
                                        
                                        
                                '//CICLO OGGETTI QUERY
                                '//------------------------------------------------------------------------------//
                                For intCiclo = 0 To .QueryDefs.count - 1
                                        
                                        'controlli
                                        '................................................................
                                            
                                
                                                Debug.Print "-----------------------------------------------------------"
                                                Debug.Print
                                                            '//stampo il nome della query + il contentuto sql
                                                            Debug.Print " stampo il NOME DELLA QUERY"
                                                            Debug.Print "  " & .QueryDefs(intCiclo).Name
                                                           
                                                            '//AGGIORNAMENTO_DEL_28_02_2020
                                                            'Stampo il contenuto sql della query
                                                            Debug.Print
                                                            Debug.Print "  " & .QueryDefs(intCiclo).SQL
                                                            
                                                            '//AGGIORNAMENTO_DEL_28_02_2020
                                                            '//stampo il tipo di query scelta direttamente
                                                            Debug.Print "IL TIPO DELLA Query"
                                                            Debug.Print "  " & .QueryDefs(intCiclo).Type
                                                
                                                Debug.Print "-----------------------------------------------------------"
                                                Debug.Print
                                                
                                            
                                        '................................................................
                                            
                                            'SALVATAGGIO NOME QUERY
                                            '................................................................
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//Aggiorno il numero oggetto query salvati nel campo della tabella _
                                                dei salvataggi del nro query.
                                                iCount = iCount + 1         '//aggiorno il contatore oggetti delle query
                                              
                                                
                                                ApriRs.AddNew
                                                '//aggiugno alla tabella il nome della query + il contenuto sql
                                                ApriRs.Fields("Name1_s") = .QueryDefs(intCiclo).Name
                                                ApriRs.Fields("IdOrd_lng") = iCount
                                                
                                                '//Tipo di oggetto caricato = QUERY
                                                ApriRs.Fields("TIPOGGETTO_s") = "QUERY"
                                                
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//SQL + TIPO DI QUERY
                                                ApriRs.Fields("Expression_m") = .QueryDefs(intCiclo).SQL
                                                            
                                                
                                                ApriRs.Update
                                                
                                                
                                                
                                                
                                                '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                        '................................................................
                                               
                                                
                                                
                                            '................................................................
                                            
                                        'controllo nome query passata con parametro
                                        If .QueryDefs(intCiclo).Name = par_sx_Name_Query Then
                                            
                                            Bln1 = True    ' bx_Flag_b True = trovato QUERY
                                            
                                            'stampo il contenuto della query scelta direttamente _
                                            solo se esiste la query ricercata.
                                            Debug.Print "  " & .QueryDefs(par_sx_Name_Query).SQL
                                            
                                            
                                            'recupero il contenuto della query
                                            RICERCA_OGGETTI_DB_PFunct = .QueryDefs(intCiclo).SQL
                                            
                                            
                                            GoTo Exit_RICERCA_OGGETTI_DB_PFunct  'esci
                                        Else
                                            Bln1 = False   ' bx_Flag_b False = non trovato QUERY
                                            RICERCA_OGGETTI_DB_PFunct = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                        End If
                                        
                                    Next intCiclo
                                    
                                '//CICLO OGGETTI QUERY  --> Next intCiclo *** fine ***
                                '//------------------------------------------------------------------------------//
                                    
                                End With
                        '//ITERAZIONE QUERYDEFS <<LE QUERY>>   *** FINE ***
                        '//====================================================================================================//
                        

            
               'CICLO NEGLI OGGETTI FORMS
               Case Is = "Cmd_Elenco_Forms"
               '//CHIAVE =====> modulo:controllo-forms
                
               
                 
   
                        '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                        '//==============================================================================================================//
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS
                        
                            
                            'Dim obj As AccessObject, obj_Oggetto As Object
                            Set obj_Oggetto = Application.CurrentProject
   
                             '//RESET DELLE VARIABILI
                            intCiclo = 0
                            iCount = 0
                           
                            'NOME form prorvvisoria ca controllare, poi cancellare
                            par_sx_Name_FORM = "form ??? da definire"

                            
                            'apro la tabella per il salvataggio delle query
                            Set ApriRs = Dbs.OpenRecordset("MSsys_FORM_Qry01_01_TUTTE")
                            
                            
                            '//RESET DELLE VARIABILI
                            intCiclo = 0
                            iCount = 0
                            
                            'pulisco la tabella delle FORMS prima del salvataggio dei nuovi nomi
                            CurrentDb.Execute ("DELETE MSsys_FORM_Qry01_01_TUTTE.* FROM MSsys_FORM_Qry01_01_TUTTE;")
                            
                            
                                                        
                        '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                        '______________________________________________________________________________
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS
                            
                                                       
                            
                        'itero nel db
                        With Dbs
                                    
                                    
                            
                        '//ITERAZIONE Forms DEFS <<LE FORMS>> - CONTROLLO FORM -
                        '//______________________________________________________________________________
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della Forms
                                
                                For Each obj In obj_Oggetto.AllForms
                                        
                                        
                                        'SALVATAGGIO NOME Forms
                                        '................................................................
                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                        '//Aggiorno il numero oggetto Forms  salvati nel campo della tabella _
                                        dei salvataggi del nro Forms.
                                          
                                          ApriRs.AddNew
                                            
                                            'CONTROLLO DEI NOMI DELE FORMS
                                            '................................................................
                                                        
                                                    '//Search for open AccessObject objects in AllForms collection. _
                                                    con la proprieta IsLoad = esegui il controllo delle Forms Aperti _
                                                    mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                    Report del Db. Per le Forms APERTI O CARICATI  viene impostata la _
                                                    proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                    FALSE. _
                                                    È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                    un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.
                                                                
                                                                
                                                                Debug.Print "  " & obj.Name
                                                                Debug.Print "  " & obj.Type
                                                                
                                                                 '//Tipo di oggetto caricato = QUERY
                                                                 ApriRs.Fields("TIPOGGETTO_s") = "FORMS"
                                                
                                               
                                                    
                                                    '//obj.Type = 2 LA FORMS
                                                    If obj.Type = 2 Then
    
                                                            If obj.IsLoaded = True Then
                                                    
                                                                   '//aggiugno alla tabella il nome della FORM solo se oggetto di tipo 2
                                                                    ApriRs.Fields("Name1_s") = obj.Name
                                            
                                                                    Debug.Print "FORM APERTA = " & obj.Name
                                                                    Debug.Print "  " & obj.Name
                                                                    Debug.Print "Nro : " & iCount
                                                                    ApriRs.Fields("PROPRIETA_s") = "True"
                                                                    ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle Forms
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
        
                                                                    
                                                            Else
                                                                    
                                                                    '//aggiugno alla tabella il nome della FORM solo se oggetto di tipo 2
                                                                    ApriRs.Fields("Name1_s") = obj.Name
                                            
                                                                    
                                                                    Debug.Print "FORM NON APERTO ---> " & obj.Name
                                                                    ApriRs.Fields("PROPRIETA_s") = "FALSE"
                                                                    ApriRs.Fields("VALOREPROPRIETA_m") = "FORM CHIUSA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle Forms
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
                                                            End If
                                                    End If
                                                              
                                                                          
                                            '................................................................
                                            
                                          ApriRs.Update
                                        '................................................................
                                      
                                Next obj
                                
                            End With

            '//ITERAZIONE FORMS DEFS <<LE FORMS>>    *** FINE ***
            '//==============================================================================================================//

                            '**************************************************************
                            
                            
                                                        
              'CICLO NEGLI OGGETTI REPORTS
               Case Is = "Cmd_Elenco_Reports"
                '//CHIAVE =====> modulo:controllo-Report
                '//Aggiornamento 22 MARZO 2020
                
                        '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                        '//______________________________________________________________________________
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS
                            
                            
                            
                            'Dim obj As AccessObject, obj_Oggetto As Object
                            Set obj_Oggetto = Application.CurrentProject
   
                            '//RESET DELLE VARIABILI
                            intCiclo = 0
                            iCount = 0
                            
                            'pulisco la tabella delle FORMS prima del salvataggio dei nuovi nomi
                            CurrentDb.Execute ("DELETE MSsys_REPORT_Qry01_01_TUTTI.* FROM MSsys_REPORT_Qry01_01_TUTTI;")
                            
                            'apro la tabella per il salvataggio dei REPORT
                            Set ApriRs = Dbs.OpenRecordset("MSsys_REPORT_Qry01_01_TUTTI")
                            
                                                        
                        '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                        '______________________________________________________________________________
                        '//OGGETTI FORMS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS
                            
                                                       
                            
                        'itero nel db
                        With Dbs
                                    
                                    
                            
                        '//ITERAZIONE REPORTS DEFS <<I REPORT>> - CONTROLLO REPORT -
                        '//______________________________________________________________________________
                        '//OGGETTI REPORTS
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della REPORTS
                                
                                For Each obj In obj_Oggetto.AllReports
                                        
                                        
                                        'SALVATAGGIO NOME REPORTS
                                        '................................................................
                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                        '//Aggiorno il numero oggetto REPORTS  salvati nel campo della tabella _
                                        dei salvataggi del nro REPORTS.
                                       
                                    
                                          
                                          ApriRs.AddNew
                                            '//aggiugno alla tabella il nome della query + il contenuto sql
                                            ApriRs.Fields("Name1_s") = obj.Name
                                            
                                            
                                            
                                            
                                            'CONTROLLO DEI NOMI DEI REPORT
                                            '................................................................
                                                        
                                                    '//Search for open AccessObject objects in AllReports collection. _
                                                    con la proprieta IsLoad = esegui il controllo dei Reports Aperti _
                                                    mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                    Report del Db. Per i reports APERTI O CARICATI  viene impostata la _
                                                    proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                    FALSE. _
                                                    È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                    un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.
                                                                
                                                                
                                                                Debug.Print "  " & obj.Name
                                                                Debug.Print "  " & obj.Type
                                                                
                                                                '//Tipo di oggetto caricato = REPORTS
                                                                ApriRs.Fields("TIPOGGETTO_s") = "REPORTS"
                                                    
                                                    '//obj.Type = 3 LA FORMS
                                                    If obj.Type = 3 Then
    
                                                            If obj.IsLoaded = True Then
                                                    
                                                                   
                                                                    Debug.Print "REPORT APERTO = " & obj.Name
                                                                    Debug.Print "  " & obj.Name
                                                                    Debug.Print "Nro : " & iCount
                                                                    ApriRs.Fields("PROPRIETA_s") = "True"
                                                                    ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle REPORTS
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
        
                                                                    
                                                            Else
                                                                    Debug.Print "REPORT NON APERTO ---> " & obj.Name
                                                                    ApriRs.Fields("PROPRIETA_s") = "FALSE"
                                                                    ApriRs.Fields("VALOREPROPRIETA_m") = "FORM CHIUSA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle REPORTS
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
                                                            End If
                                                    End If
                                                               
                                            '................................................................
                                            
                                            
                                            
                                            
                                          
                                          ApriRs.Update
                                        '................................................................
                                        
                                        
                                        
                                    
                                Next obj
                                
                            End With

   
                 
                        '//ITERAZIONE REPORTS DEFS <<I REPORT>>
                        '//______________________________________________________________________________
                            
                            
               'CICLO NEGLI OGGETTI MACRO - CONTROLLO MACRO -
               Case Is = "Cmd_Elenco_MACRO"
                '//CHIAVE =====> modulo:controllo-Macro
                '//Aggiornamento 22 MARZO 2020
                
                        '//ITERAZIONE MACRO DEFS <<LE MACRO>>
                        '//______________________________________________________________________________
                        '//OGGETTI MACRO
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO
                            
                            
                            
                            'Dim obj As AccessObject, obj_Oggetto As Object
                            Set obj_Oggetto = Application.CurrentProject
   
                            '//RESET DELLE VARIABILI
                            intCiclo = 0
                            iCount = 0
                            
                            'pulisco la tabella delLE MACRO prima del salvataggio dei nuovi nomi
                            CurrentDb.Execute ("DELETE MSsys_MACRO_Qry01_01_TUTTE.* FROM MSsys_MACRO_Qry01_01_TUTTE;")
                            
                            'apro la tabella per il salvataggio dei Macro
                            Set ApriRs = Dbs.OpenRecordset("MSsys_MACRO_Qry01_01_TUTTE")
                            
                                                        
                        '//ITERAZIONE MACRO DEFS <<LE MACRO>>
                        '______________________________________________________________________________
                        '//OGGETTI MACRO
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO
                            
                                                       
                            
                        'itero nel db
                        With Dbs
                                    
                                    
                            
                        '//ITERAZIONE MACRO DEFS <<I Macro>>
                        '//______________________________________________________________________________
                        '//OGGETTI MACRO
                        '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO
                                
                                For Each obj In obj_Oggetto.AllMacros
                                        
                                        
                                        'SALVATAGGIO NOME MACRO
                                        '................................................................
                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                        '//Aggiorno il numero oggetto MACRO  salvati nel campo della tabella _
                                        dei salvataggi del nro MACRO.
                                       
                                    
                                          
                                          ApriRs.AddNew
                                            '//aggiugno alla tabella il nome della query + il contenuto sql
                                            ApriRs.Fields("Name1_s") = obj.Name
                                            
                                                    
                                            
                                            
                                            
                                            'CONTROLLO DEI NOMI DEI Macro
                                            '................................................................
                                                        
                                                    '//Search for open AccessObject objects in AllMACRO collection. _
                                                    con la proprieta IsLoad = esegui il controllo dei MACRO Aperti _
                                                    mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                    Macro del Db. Per i MACRO APERTI O CARICATI  viene impostata la _
                                                    proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                    FALSE. _
                                                    È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                    un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.
                                                                
                                                                
                                                                Debug.Print "  " & obj.Name
                                                                Debug.Print "  " & obj.Type
                                                                
                                                                '//Tipo di oggetto caricato = MACRO
                                                                ApriRs.Fields("TIPOGGETTO_s") = "MACRO"
                                                
                                                    
                                                    '//obj.Type = 4 LA MACRO
                                                    If obj.Type = 4 Then
    
                                                           
                                                    
                                                                   
                                                                    Debug.Print "Macro APERTO = " & obj.Name
                                                                    Debug.Print "  " & obj.Name
                                                                    Debug.Print "Nro : " & iCount
                                                                    'ApriRs.Fields("PROPRIETA_s") = "True"
                                                                    'ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"
                                                                    
                                                                     iCount = iCount + 1         '//aggiorno il contatore oggetti delle MACRO
                                                                     Debug.Print "Nro : " & iCount
                                                                     
                                                                     '//NRO di oggetto
                                                                     ApriRs.Fields("IdOrd_lng") = iCount
        
        
                                                                    
                                                        
                                                    End If
                                                               
                                            '................................................................
                                            
                                            
                                            
                                            
                                          
                                          ApriRs.Update
                                        '................................................................
                                        
                                        
                                        
                                    
                                Next obj
                                
                            End With

   
                 
                        '//ITERAZIONE MACRO DEFS <<I Macro>>
                        '//______________________________________________________________________________
         
                            
                    
                    Case Else
                            MsgBox "Scelta comando non effettuata", vbExclamation
                             GoTo Exit_RICERCA_OGGETTI_DB_PFunct
                End Select
                            'chiudo i rs aperti + Rilascio la memoria
                            ApriRs.Close
                            Set ApriRs = Nothing
                            
                            
                                    
'------------------------------------------------------------------------------
'                       FINE FUNCTIONE E GESTIONE ERRORI

Exit_RICERCA_OGGETTI_DB_PFunct:
    Dbs.Close
    Set Dbs = Nothing
    Exit Function


Err_RICERCA_OGGETTI_DB_PFunct:
    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.number & " - " & Err.Description, vbCritical, "RICERCA_OGGETTI_DB_PFunct"
    Resume Exit_RICERCA_OGGETTI_DB_PFunct
    
End Function

'//*** FINE ***
'//                                     SALVATAGGIO OGGETTI DATABASE
'//{@=============================================================================================================@}




'//###########################################################################################################//
'//####
'//####                             PROCEDURE DI RICERCA   *** FINE ***
'//####
'//###########################################################################################################//



'// *** FINE ***
'//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Salvataggio_ATTIVITA
'//SALVATAGGIO PARAMETRI NELLA TABELLA TMP
'//###########################################################################################################################//








'//###########################################################################################################//
'//####
'//####                             PROCEDURE DI CONTROLLO *** INIZIO ***
'//#### CONTROLLO/ CONTROLLA...
'//###########################################################################################################//


        '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
        '//=======================================================================================================//
        '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva
        '//note: è stata inserita nella classe CLASSE_GESTIONE_OGGETTI_FORM
        Public Function Controlla_SE_OGGETTO_FORM_ESISTE_pFunct(par_sxNameForm As String) As Boolean
            
            
            
            
            
            
            '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
            '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva.LaClasse
            '//---------------------------------------------------------------------------------//
            '//NOTA  :Controlla se esiste la form negli oggetti del database
                
            
            
                        'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                        'Assegno e creo l'oggetto di classe
                        'Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As Classe_GESTIONE_OGGETTI_FORM
                        'Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New Classe_GESTIONE_OGGETTI_FORM
                        
                        
                        'PARAMETRI DI ESTRAZIONE
                        'Attivo il metodo della classe rappresentata dalla routine passando
                        'sxNameForm = "Nome della form"
                        'MyCls_CLASSE_GESTIONE_OGGETTI_FORM.Controlla_SE_OGGETTO_FORM_ESISTE_pFunct sxNameForm
                        
                        'Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing
                                                    
                '//---------------------------------------------------------------------------------//
            
            
        
            
            
            
            
             
            Dim formName_s As String
            Dim ctl As Control
                
            '//reset
            FormGestione_b = False
            
                
            'application.Form[AA_GEST_FrmM00_GESTIONE_MASTER].
            Dim obj As AccessObject, Dbs As Object
            Set Dbs = Application.CurrentProject
            ' Search for open AccessObject objects in AllForms collection.
            For Each obj In Dbs.AllForms
               
                    ' Print name of obj.
                    Debug.Print obj.Name
                    formName_s = obj.Name
                     
                    '//CONTROLLO SUL NOME
                    If formName_s = par_sxNameForm Then
                        
                       '//La form esiste restituisco True
                       FormGestione_b = True
                       Controlla_SE_OGGETTO_FORM_ESISTE_pFunct = True
                    
                    Else
                         FormGestione_b = False
                       
                    End If
                    
                
            Next obj
            
            
                    If Bool1 = False Then
                        MsgBox "ATTENZIONE LA FORM ---> " & formName_s & " NON ESISTE NEL DB CORRENTE", vbExclamation, "CONTROLLO FORM"
                    End If
        
        
            
                  
            
                Set obj = Nothing
                Set Dbs = Nothing
                Set ctl = Nothing
                formName_s = vbNull
                
            
            
        
        End Function
        '//*** FINE ***
        '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
        '//=======================================================================================================//
        
        
        
        
        '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DEL_SINGOLO_OGGETTO_NEL_DB_CORRENTE
        '//=======================================================================================================//
        '//CODICE----->ACONTROLLA_OGGETTO_pFunction.Funcition
        
        '//ATTIVO LA FUNZIONE TROVA L'OGGETTO NEL DB CORRENTE
        '//-----------------------------------------------------------------------------------//
        '//NOTE : attivo la ricerca dell'oggetto nel DB corrente iterando in tutti gli oggetti _
                se trova l'oggetto passato come parametro, tipo form, modulo, query od altro _
                ritrona il parametro OGGETTO TROVATO con il quale viene effettuato un controllo _
                con la if che da il permesso di attivare la procedura della classe. Se il parametro _
                ritorna NON TROVATO l'attivà della classe viene saltata esplicitamente.
        '//
        '//CODICE----->ACONTROLLA_OGGETTO_pFunction.Funcition.ATTIVA
            'Dim ObjDaTrovare_s As String
            'ObjDaTrovare_s = ""
            'Str1 = CONTROLLA_OGGETTO_pFunction(ObjDaTrovare_s)
                'If ObjDaTrovare_s = "OGGETTO TROVATO" Then
                    'eseguo l'attività della classe altrimenti non viene eseguita
                    '...
                'End If
        
        '//Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct = CONTROLLO ESISTENZA OGGETTO TABELLA
        '//-----------------------------------------------------------------------------------//
        Public Function Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct(par_sxNameTABELLA As String)
        
            
            Dim obj As AccessObject
            Dim Dbs As Object
        
            Dim objAllTables As AllTables
            
              '//Variabile database corrente + reset
            Set Dbs = Application.CurrentData
            Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct = False            '//imposto a ricerca fallita
          
            
            'Search for open AccessObject objects in AllForms collection.
            For Each obj In Dbs.AllTables
        
                    ' Print name of obj.
                    Debug.Print obj.Name
                    TableName_s = obj.Name
                     
                    '//CONTROLLO SUL NOME
                    If TableName_s = par_sxNameTABELLA Then
                        
                       '//Se esiste la TABELLA restituisco True
                     
                       Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct = True
                       
                            '//controllo TABELLE ESISTENTE
                             Debug.Print "----------------------LA TABELLA XLS ESISTE --------------------------------------"
                             Debug.Print
                             Debug.Print "tabella xls passata con parametro par_sxNameTABELLA ESISTE nel db               ----> " & par_sxNameTABELLA
                            
                       
                       Exit For
                       
                    End If
                    
             Next obj
            
        
        End Function
        
        '//*** FINE ***
        '//Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct = CONTROLLO ESISTENZA OGGETTO TABELLA
        '//-----------------------------------------------------------------------------------//
        
        
        '//Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct = CONTROLLO ESISTENZA OGGETTO QUERY
        '//-----------------------------------------------------------------------------------//
        Public Function Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(par_sxNameQuery As String)
        
            
            Dim obj As AccessObject
            Dim Dbs As Object
        
            Dim objAllQueries As AllQueries
            
              '//Variabile database corrente + reset
            Set Dbs = Application.CurrentData
            Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct = False            '//imposto a ricerca fallita
          
            
            'Search for open AccessObject objects in AllForms collection.
            For Each obj In Dbs.AllQueries
        
                    ' Print name of obj.
                    Debug.Print obj.Name
                    formName_s = obj.Name
                     
                    '//CONTROLLO SUL NOME
                    If formName_s = par_sxNameQuery Then
                        
                       '//Se esiste la query restituisco True
                     
                       Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct = True
                      
                       Exit For
                       
                    End If
                    
             Next obj
            
        
        End Function
        
        '//*** FINE ***
        '//Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct = CONTROLLO ESISTENZA OGGETTO QUERY
        '//-----------------------------------------------------------------------------------//
        
        Public Function CONTROLLA_OGGETTO_pFunction(ByVal NameObject_s As String) As String
        
        
        Dim obj As Object
        Dim frm As Form
            Dim Ctrl As Control
        
            '//impsto il valore a non trovato
            CONTROLLA_OGGETTO_pFunction = "NON TROVATO"
            
            For Each obj In CurrentProject.AllForms
                If InStr(obj.Name, "MODELLO_Frm01_M01_GE_FILTRO_M15_H15_L35_NoButton") Then Debug.Print "Forms|" & obj.Name
                    
                    'Trovato l'oggetto nel db corrente
                    CONTROLLA_OGGETTO_pFunction = "TROVATO"
                
                'Set frm = Forms(obj.Name) '- this had problems as well, so I commented out
                For Each Ctrl In forms(obj.Name).Controls
                    If InStr(Ctrl.Name, "LBL_FILTRO_01") Then Debug.Print "Form|" & obj.Name & "|" & Ctrl.Name
                Next Ctrl
            Next obj
            
        End Function
        
        
        
        '//*** FINE ***
        '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DEL_SINGOLO_OGGETTO_NEL_DB_CORRENTE
        '//=======================================================================================================//



        '//CODICE-->Controlla_FormGESTIONE_APERTA.01
        '//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA.01
        '//=======================================================================================================//
        '//SI TROVA NEL MODULO -> AA_GEST_FrmM00_GESTIONE_MASTER_Mdl01_ControlloFormAperta
        
        Public Function Controlla_FormGESTIONE_APERTA() As Boolean
            
            
        On Error GoTo Controlla_FormGESTIONE_APERTA_Err
        
             
            
            '//MODELLO ISTANZIAMENTO CLASSE ATTI DI GESTIONE
            '//.....................................................................................................
              
               '// CONTROLLA SE LA FORM GESTIONE E' APERTA
               '//CODICE-->Controlla_FormGESTIONE_APERTA.01.CALL
               '//-------------------------------------------------------------------------//
                   '//STR1 = Controlla_FormGESTIONE_APERTA()
               '//-------------------------------------------------------------------------//
            
                '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                    'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione
            
                '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
                   l'oggetto di classe alla variabile e creo creo l'oggetto di classe
                    'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione
                
                '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
                    'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
                    'MyCls_CLASSE_ATTI_Tb01_Gestione.APRO_PATH_O_FILE_GENERICI_pSub (Str1)
                    
                '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                    'Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing
            
            
            
            '//.....................................................................................................
    


'//CHIAMATA DELLA FUNZIONE E FUNZIONE PER IL CONTROLLO ESISTENZA FORM
'//############################################################################################################################//

                 '// CONTROLLA SE LA FORM GESTIONE E' APERTA
                 '//-------------------------------------------------------------------------//
                 '//Da inserire nell'evento per il controllo della form aperta. Chiama la _
                  routine che attiva la classe.
                  
                         '//CHIAMO LA FUNZIONE
                         '//--> FormGestione_b = CONTROLLO_FORM_APERTA
                         
              
                 '//-------------------------------------------------------------------------//
            
            
            '//INSERIRE QUESTA FUNZIONE
            '//--> Private Function CONTROLLO_FORM_APERTA() As Boolean
            
                '//RESET
                '//--> CONTROLLO_FORM_APERTA = False
            
                        '//MODELLO ISTANZIAMENTO CLASSE ATTI DI GESTIONE
                        '//.....................................................................................................
                            
                        
                            '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                                '//--> Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione
                        
                            '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
                               l'oggetto di classe alla variabile e creo creo l'oggetto di classe
                                '//--> Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione
                            
                            '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
                                'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
                                '//--> FormGestione_b = MyCls_CLASSE_ATTI_Tb01_Gestione.Controlla_FormGESTIONE_APERTA()
                                '//--> CONTROLLO_FORM_APERTA = FormGestione_b
                                
                                
                            '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                                '//--> Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing
                        
                        '//.....................................................................................................
                
            
            
            '//--> End Function
            

'//*** FINE ***
'//CHIAMATA DELLA FUNZIONE E FUNZIONE PER IL CONTROLLO ESISTENZA FORM
'//############################################################################################################################//


            
            
        '//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA.01_Esegui
        '//__________________________________________________________________________________________//
        
            Dim formName_s As String
            Dim ctl As Control
                
            '//reset
            FormGestione_b = False
        
            
            'application.Form[AA_GEST_FrmM00_GESTIONE_MASTER].
            Dim obj As AccessObject, Dbs As Object
            Set Dbs = Application.CurrentProject
            ' Search for open AccessObject objects in AllForms collection.
            For Each obj In Dbs.AllForms
                                
                If obj.IsLoaded = True Then
                    ' Print name of obj.
                    Debug.Print obj.Name
                    formName_s = obj.Name
                     
                    '//CONTROLLO SUL NOME
                    If formName_s = "AA_GEST_FrmM00_GESTIONE_MASTER" Then
                        
                       '//FormGestione_b = true per indicare che la form è aperta e restituisco il nome della form _
                       alla routine chiamante
                       FormGestione_b = True
                       Controlla_FormGESTIONE_APERTA = True
                       
                    End If
                    
                End If
            Next obj
            
                Set obj = Nothing
                Set Dbs = Nothing
                Set ctl = Nothing
                formName_s = vbNull
            
        '//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA.01_Esegui *** fine ***
        '//__________________________________________________________________________________________//
                
                
                
                
        '// USCITA E CONTROLLO ERRORI
        '//__________________________________________________________________________________________//
                
            
Controlla_FormGESTIONE_APERTA_Exit:
            Exit Function
        
Controlla_FormGESTIONE_APERTA_Err:
            MsgBox Error$
            Resume Controlla_FormGESTIONE_APERTA_Exit
            
        
        End Function
        '//*** FINE ***
        '//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA
        '//=======================================================================================================//
        
        


'//###########################################################################################################//
'//####
'//####                             PROCEDURE DI CONTROLLO *** FINE ***
'//#### CONTROLLO/ CONTROLLA...
'//###########################################################################################################//







'//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@//
'//                                             ESPORTAZIONI DATI PER IL SETTORE LAVORI PUBBLICI
'//
'//NOTE :ESPORTA LE TABELE XLS ESISTENTI        14.08.2020 del settore lavori pubblici.

'//      ATTIVA IN VISUALIZZAZIONE LE QUERY ESISTENTI
'//     TODO: non gestisce le tabelle collegate, da errore se esistono. Per il momento le tabelle collegate non funzionanti _
              sono state cancellate. _
              GESTIRE LE TABELLE COLLEGATE!!!!
'//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@//




'//             AGGIORNA I DATI CON LE QUERY ED ESPORTA IN MDB + XLS
'//****************************************************************************************************************//
'------------------------------------------------------------
' CODICE----->AggiornaDati_ed_EsportaDati_Pfunct.ATTIVA
' AggiornaDati_ed_EsportaDati_Pfunct
' si trova in questo modulo                 ----> LLPP_ATTI_Mdl01_AGGIORNA_Mcr01_01_AggiornaDati+EsportaDati
' ma sistemato all'interno della classe     ----> CLASSE_ATTI_Tb01_Gestione _
' perchè all'interno della classe posso fare un controllo degli oggetti prima di _
' attivare il comando Cmd; se non esistono nel db corrente da errore; ma _
' con il controllo degli oggetti puoi attivare quelli che sono presenti.
'------------------------------------------------------------
Public Function AggiornaDati_ed_EsportaDati_Pfunct()

'//Variabile per le tabelle ESPORTATE E NON ESPORTATE
Dim MsgTabelle_ESPORTATE_s As String
Dim iCount_TABELLE_ESPORTATE As Integer
Dim MsgTabelle_NON_ESPORTATE_s As String
Dim iCount_NON_TABELLE_ESPORTATE As Integer


Dim OggettoDaRicercare_s   As String
Dim iCount As Integer

On Error GoTo AggiornaDati_ed_EsportaDati_Pfunct_Err

    
    
    
    '//RESET
    MsgTabelle_NON_ESPORTATE_s = ""
    iCount = 0
    
' AGGIORNAMENTI + ACCODAMENTI
' ***********************************************************************************************************************
' Note  : Eseguo l'aggiornamento dei vari argomenti di controllo.
    
    
    '//PRES3000 AGGIORNAMENTI TABELLE PRESENZE MEDIANTE LA MACRO
    '//=========================================================================================================
        
              
        
        '//RIRCERCA OGGETTO NEL DB CORRENTE - Query da ricercare
        '//.......................................................................................
        '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
            
            '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
               di ricerca secondo il tipo di oggetto
               
            Bool1 = False
            OggettoDaRicercare_s = "PRES3000_Qry01_10_AGGIORNA_Campo_IDKEY_CALENDARIO"      '//Aggiorno ID nel campo idKey del calendario, per preservare la chiave univoca.
            Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                
                '//SE TRUE = oggetto esistente ed eseguo il comando
                If Bool1 Then
                    '//!!!! ATTENZIONE NON GESTISCE LE TABELLE COLLEGATE SISTEMARE ????????
                    
                    '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                    acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                    acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                    DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                        
                        '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                        iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                        MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                Else
                        
                        '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                        iCount = iCount_NON_TABELLE_ESPORTATE + 1
                        MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                End If
                
              
       '//.......................................................................................


       
        '//*** AGGIORNO LE TIMBRATURE ***
        '//RIRCERCA OGGETTO NEL DB CORRENTE - Query da ricercare
        '//.......................................................................................
        '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
            
            '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
               di ricerca secondo il tipo di oggetto
               
            Bool1 = False
            OggettoDaRicercare_s = "PRES3000_Qry01_11_AGGIORNA_Campo_IDKEY_GIORNALIERE"  '//Aggiorno ID nel campo timbrature giornaliere
            Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                
                '//SE TRUE = oggetto esistente ed eseguo il comando
                If Bool1 Then
                    
                      '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                    acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                    acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                       DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                        
                        '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                        iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                        MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                Else
                        
                        '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                        iCount = iCount_NON_TABELLE_ESPORTATE + 1
                        MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                End If
                
          
       '//.......................................................................................

        
    '//PRES3000 AGGIORNAMENTI TABELLE PRESENZE MEDIANTE LA MACRO*** FINE ***
    '//=========================================================================================================
    


    
    '//AGGIORNAMENTI
    '//=========================================================================================================
    
        
        '//INTESTATARI Aggiorna tabella INTESTARI - campo ->Argomento
        '//================================================================================================================================//
               
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "AGGIORNA_INTESTATARI_QyN01-01_Campo_Argomento"    '//AGGIORNA = la tabella INTESTARI - campo ->Argomento
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
                        
                        
                        DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                        
                        '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                        iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                        MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                Else
                        
                        '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                        iCount = iCount_NON_TABELLE_ESPORTATE + 1
                        MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                End If
                
                    
                        
               '//.......................................................................................
        

       '//================================================================================================================================//

    
        '//TABELLA ATTI AGGIORNAMENTI
        '//************************************************************************************************************************************//
            
                    
                '//01) Aggiorna tabella TABELLA ATTI - campo ->IDKeyArch+IdOrd e contestualmente aggiorno la descrizione opera - campo ->DescrCodOpera_s
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_ATTI_Qy01-01_Gestione_AGGIORNA_IDKeyArch+IdOrd"    '//AGGIORNA = campi ->IDKeyArch+IdOrd+DescrCodOpera_s
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                        
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
               '//.......................................................................................
        
      
            
                        
                '//02) Aggiorna tabella TABELLA ATTI - campo ->DenomFascicolo_s e la denominazione del fascicolo prendendo i dati dalla tabella [LPP_DF06_FASCICOLI].[DenomFascicolo_s]
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFascicolo"    '//AGGIORNA = campi ->DenomFascicolo_s CON IL CAMPO [LPP_DF06_FASCICOLI].[DenomFascicolo_s]
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                                 DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
               '//.......................................................................................
       
                
                                
                '//03) Aggiorna tabella TABELLA ATTI - campo ->ANNO l'anno di creazione del documento viene aggiornato solo se null
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_ATTI_Qy01-11_AGGIORNA_ANNO"    '//AGGIORNA = campo ->ANNO se > null
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
               '//.......................................................................................
        
            
            
        '//ATTI *** fine ***
        '//************************************************************************************************************************************//
            
            
            
          
        '//FALDONI + FASCICOLI *** fine ***
        '//************************************************************************************************************************************//
            
               
                
                                
                '//01) Aggiorna tabella TABELLA ATTI - campo ->->CodDenomFaldone_s + DenomFaldone_s + AnnoFaldone_i e la denominazione DEL FALDONE
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFALDONE"    '//AGGIORNA = campi ->CodDenomFaldone_s + DenomFaldone_s + AnnoFaldone_i
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
               '//.......................................................................................
        
            
            
            
                       
                '//02) Aggiorna tabella TABELLA ATTI - campo ->CodDenomFaldone_s (la denominazione DEL FASCICOLO)
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFascicolo"    '//AGGIORNA = campo ->CodDenomFaldone_s
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
                '//.......................................................................................
        
            
            
                    
                    
            
            
          
        '//FALDONI + FASCICOLI *** fine ***
        '//************************************************************************************************************************************//
              
            
               
        
        
        '//ALLEGATI
        '//'//************************************************************************************************************************************//
            
            
                        
                '//01) Aggiorna tabella TABELLA ATTI - campo ->CodOpera_s
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_ATTI_Qy02_01_Allegati_AGGIORNA_CodOpera"    '//AGGIORNA = campo ->CodOpera_s
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
                '//.......................................................................................
        
                        
                '//02) Aggiorna tabella TABELLA ATTI - i campi ->IDGestione + AnnoAtto_i + CodPratica_s +NroPratica_i
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_ATTI_Qry02_01_Allegati_AGGIORNA_IDGestione"    '//AGGIORNA = campi ->IDGestione + AnnoAtto_i + CodPratica_s +NroPratica_i
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                     
                '//.......................................................................................

            
                '//03) Aggiorna tabella TABELLA ATTI - i campi ->NroPratica (il campo viene aggiornato con 0 (zero) solo se nullo)
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_ATTI_Qy01_01_Gestione_AGGIORNA_Campo_NroPratica_s=Null"    '//AGGIORNA = campi ->NroPratica
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
                '//.......................................................................................
    
            
            
       '//************************************************************************************************************************************//
            
            
        '//AFFIDAMENTI
        '//************************************************************************************************************************************//

            
                '//01) Aggiorna tabella TABELLA ATTI - i campi -> I dati degli affidamenti
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "AGGIORNA_AFFIDAMENTI_QyN10-01_Campi_TipoAtto+NroAtto+DataAtto"    '//AGGIORNA = campi -> AFFIDAMENTI vari campi
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
                '//.......................................................................................

        '//************************************************************************************************************************************//
    
    
    
        '//ACCODAMENTI
        '//************************************************************************************************************************************//
    
  
        
                '//01) Aggiorna tabella TABELLA ATTI - i campi -> I dati degli affidamenti
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "Accoda_QyDF15-01_NwEVENTI_Ins"    '//ACCODO = campi -> Nuovi eventi
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
                '//.......................................................................................


        
                    
                '//02) ACCODO I NUOVI CODICI - LLPP_QyDF01_CodiceOpera_ACCODA_CODICI_Ins
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "LLPP_QryDF01_CodiceOpera_ACCODA_CODICI_Ins"    '//ACCODO = nuovi codice opera dalla tabelal GESTIONE campi -> Codici Opera
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
                '//.......................................................................................

        '//*** FINE ***
        '//ACCODAMENTI
        '//************************************************************************************************************************************//
                         
          
        
    
    '//TIMBRATURE - PRES3000
    '//************************************************************************************************************************************//
    '//NOTE : accodo timbrature ed orari nella tabella PRES3000
    
                                  
                '//01) ACCODO NUOVE ENTRATE  nella tabella orari.
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "PRES3000_Qry05_N01_ACCODA_Nw_ENTRATE->Orari"    '//ACCODO = I NUOVI ORARI
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
                                
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                '//.......................................................................................

    
                '//02) ACCODO NUOVE ENTRATE  nella tabella orari.
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    OggettoDaRicercare_s = "PRES3000_Qry05_N02_ACCODA_Nw_USCITE->Orari"    '//ACCODO = LE NUOVE USCITE
                    Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)
                        
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                            acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
        
                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1
                                
                                '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        Else
                                
                                '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                iCount = iCount_NON_TABELLE_ESPORTATE + 1
                                MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                        End If
                        
                '//.......................................................................................

    
        
 
    
    
    '//*** FINE ***
    '//TIMBRATURE
    '//************************************************************************************************************************************//

' ***********************************************************************************************************************


'ESPORTAZIONI
'***********************************************************************************************************************
    
    '//TABELLA GESTIONE
    '//============================================================================================================
       
       
       '//ESPORTO ATTI DI GESTIONE
       '//Esporto la tabella con la specifica
       '//ESPORTA I DATI ---------------------------------------------------------------------------------------------------------------------------------------> False = non visualizza il file
       '//SOSPESA
       'DoCmd.TransferText acExportDelim, "LLPP_ATTI_Tb01_Gestione - specifica di esportazione", "LLPP_ATTI_Tb01_Gestione", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ATTI_Tb01_Gestione.txt"
       'Esporto tabella in dos
       'DoCmd.OutputTo acOutputTable, "LLPP_ATTI_Tb01_Gestione", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ATTI_Tb01_Gestione_DOS.txt", False, "", 0, acExportQualityPrint
       '//Apro la tabella esportata
       'DoCmd.OpenTable "LLPP_ATTI_Tb01_Gestione", acViewNormal, acEdit
                      
       
       
       '//TABELLA ALLEGATI
       '//Esporto la tabella con la specifica - LLPP_ATTI_Tb02_Allegati - specifica di esportazione
       '//SOSPESA
       'DoCmd.TransferText acExportDelim, "LLPP_ATTI_Tb02_Allegati - specifica di esportazione", "LLPP_ATTI_Tb02_Allegati", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ATTI_Tb02_Allegati.txt"
       '//Esporto tabella in dos
       'DoCmd.OutputTo acOutputTable, "LLPP_ATTI_Tb02_Allegati", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ATTI_Tb02_Allegati_Dos.txt", False, "", 0, acExportQualityPrint
       '//Apro la tabella esportata
       'DoCmd.OpenTable "LLPP_ATTI_Tb02_Allegati", acViewNormal, acEdit
       
       
       '//GLI INDIRIZZI
       '//........................................................................................................
            '//TABELLA INTESTATARI
            '//Esporto la tabella con la specifica - INDIRIZZI_Tb01_INTESTATARI - specifica di esportazione
            '//SOSPESA
            'DoCmd.TransferText acExportDelim, "INDIRIZZI_Tb01_INTESTATARI - specifica di esportazione", "Indirizzi_Tb01_INTESTATARI", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\INDIRIZZI_Tb01_INTESTATARI.txt"
            '//Esporto tabella in dos
           ' DoCmd.OutputTo acOutputTable, "Indirizzi_Tb01_INTESTATARI", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\INDIRIZZI_Tb01_INTESTATARI_Dos.txt", False, "", 0, acExportQualityPrint
            '//Apro la tabella esportata
            'DoCmd.OpenTable "Indirizzi_Tb01_INTESTATARI", acViewNormal, acEdit
            
         
        
         
            '//TABELLA INDIRIZZI ELENCO
            '//Esporto la tabella con la specifica - INDIRIZZI_Tb02_ELENCO - specifica di esportazione
            '//SOSPESA
            'DoCmd.TransferText acExportDelim, "INDIRIZZI_Tb02_ELENCO - specifica di esportazione", "Indirizzi_Tb02_ELENCO", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\INDIRIZZI_INDIRIZZI_Tb02_ELENCO.txt"
            '//Esporto tabella in dos
            'DoCmd.OutputTo acOutputTable, "Indirizzi_Tb02_ELENCO", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\INDIRIZZI_INDIRIZZI_Tb02_ELENCO_Dos.txt", False, "", 0, acExportQualityPrint
            '//Apro la tabella esportata
            'DoCmd.OpenTable "Indirizzi_Tb02_ELENCO", acViewNormal, acEdit
            
          '//........................................................................................................
       
     
    'Indirizzi_Tb02_ELENCO - specifica di esportazione
    
    '//TABELLA GESTIONE *** FINE ***
    '//============================================================================================================
     
    '//PRES3000
    '//============================================================================================================
        
        '//IL CALENDARIO
        'Esporto la tabella con la specifica
        'ESPORTA I DATI ---------------------------------------------------------------------------------------------------------------------------------------> False = non visualizza il file
        '//SOSPESA
        'DoCmd.TransferText acExportDelim, "PRES3000_Tb01_Calendario - specifica di esportazione", "PRES3000_Tb01_Calendario", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_Tb01_Calendario.txt"
        'Esporto tabella in dos
        'DoCmd.OutputTo acOutputTable, "PRES3000_Tb01_Calendario", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_Tb01_Calendario_DOS.txt", False, "", 0, acExportQualityPrint
        '//Apro la tabella esportata
        'DoCmd.OpenTable "PRES3000_Tb01_Calendario", acViewNormal, acEdit
        
        
        '//LE GIORNALIERE - PRES3000_Tb02_ElencoGiornaliere
        'Esporto la tabella con la specifica
        'ESPORTA I DATI ---------------------------------------------------------------------------------------------------------------------------------------> False = non visualizza il file
        '//SOSPESA
        'DoCmd.TransferText acExportDelim, "PRES3000_Tb02_ElencoGiornaliere - specifica di esportazione", "PRES3000_Tb02_ElencoGiornaliere", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_Tb02_ElencoGiornaliere.txt"
        'Esporto tabella in dos
        'DoCmd.OutputTo acOutputTable, "PRES3000_Tb02_ElencoGiornaliere", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_Tb02_ElencoGiornaliere_DOS.txt", False, "", 0, acExportQualityPrint
        '//Apro la tabella esportata
        'DoCmd.OpenTable "PRES3000_Tb02_ElencoGiornaliere", acViewNormal, acEdit
        
        
        
        '//LA DEFINIZIONI
        
        'LLPP_DF01_CodiceOpera
        
        'ESPORTA I DATI ---------------------------------------------------------------------------------------------------------------------------------------> False = non visualizza il file
        '//SOSPESO
        'DoCmd.TransferText acExportDelim, "LLPP_DF01_CodiceOpera - specifica di esportazione", "LLPP_DF01_CodiceOpera", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_DF01_CodiceOpera.txt"
        'Esporto tabella in dos
       ' DoCmd.OutputTo acOutputTable, "LLPP_DF01_CodiceOpera", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_DF01_CodiceOpera_DOS.txt", False, "", 0, acExportQualityPrint
        '//Apro la tabella esportata
        'DoCmd.OpenTable "LLPP_DF01_CodiceOpera", acViewNormal, acEdit
        
    
    
        
    '//PRES3000
    '//============================================================================================================


    '//ESPORTA_I_FILE_IN_FORMATO_XLS
    '//============================================================================================================

                
                '//ATTIVO LA PROCEDURA DI ESPORTAZIONE DEI FILE IN FORMATO XLS DEGLI ATTI DI IMPEGNO
                '//.................................................................................................
                '//NOTE: Dim + Reset ed impostazioni
                
                '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls
                'ESPORTA_FILE_XLS_ConMatriceDati_pFunct "", _
                "LLPP_ATTI_Tb01_Gestione", _
                "LLPP_ATTI_Tb02_Allegati", _
                "Indirizzi_Tb01_INTESTATARI", _
                "Indirizzi_Tb02_ELENCO", _
                "PRES3000_Tb01_Calendario", _
                "PRES3000_Tb02_ElencoGiornaliere", _
                "LLPP_DF01_CodiceOpera", _
                "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", _
                "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"

'//ESPORTA I DATI IN FILE XLS
'//........................................................
 


   '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls _
      ATTENZIONE, occorre inserire un record vuoto "", affinche possono essere considerati tutti gli altri _
      file dell'esportazione altrimenti di viene saltato il primo "LLPP_ATTI_Tb01_Gestione", _
      PER ESPORTARE CHIAMARE QUESTA FUNZIONE : CODICE------> ESPORTA_FILE_XLS_ConMatriceDati_pFunct
      
                ESPORTA_FILE_XLS_ConMatriceDati_pFunct "", _
                "LLPP_ATTI_Tb01_Gestione", _
                "LLPP_ATTI_Tb02_Allegati", _
                "Indirizzi_Tb01_INTESTATARI", _
                "Indirizzi_Tb02_ELENCO", _
                "PRES3000_Tb01_Calendario", _
                "PRES3000_Tb02_ElencoGiornaliere", _
                "LLPP_DF01_CodiceOpera", _
                "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", _
                "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"

'//........................................................

              
                
               

                    
                 '//APRO LE  TABELLE ESPORTATE  --- SOSPESO
                 'DoCmd.OpenTable "LLPP_ATTI_Tb01_Gestione", acViewNormal, acEdit
                 'DoCmd.OpenTable "LLPP_ATTI_Tb02_Allegati", acViewNormal, acEdit
                 'DoCmd.OpenTable "Indirizzi_Tb01_INTESTATARI", acViewNormal, acEdit
                 'DoCmd.OpenTable "Indirizzi_Tb02_ELENCO", acViewNormal, acEdit
                 'DoCmd.OpenTable "PRES3000_Tb01_Calendario", acViewNormal, acEdit
                 'DoCmd.OpenTable "PRES3000_Tb02_ElencoGiornaliere", acViewNormal, acEdit
                 'DoCmd.OpenTable "LLPP_DF01_CodiceOpera", acViewNormal, acEdit
                 'DoCmd.OpenTable "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", acViewNormal, acEdit
                 'DoCmd.OpenTable "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA", acViewNormal, acEdit
                 
                 
                '//.................................................................................................

    '//ESPORTA_I_FILE_IN_FORMATO_XLS         *** FINE ***
    '//============================================================================================================
     
     
'***********************************************************************************************************************
    
                            
            If MsgTabelle_ESPORTATE_s > "" Then
                        
                MsgBox "LE TABELLE ESPORTATE SONO LE SEGUENTI: --->  " & MsgTabelle_ESPORTATE_s, vbExclamation
            End If
                            
                            
            If MsgTabelle_NON_ESPORTATE_s > "" Then
                        
                MsgBox "ATTENZIONE QUESTI OGGETTI NON SONO STATI TROVATI --->  " & MsgTabelle_NON_ESPORTATE_s, vbExclamation
            End If
            

    MsgBox "FINE ATTIVITA' DI Esportazione, aggiornamento ed accodamento. Operazioni effettuate!", vbExclamation
   

AggiornaDati_ed_EsportaDati_Pfunct_Exit:
    Exit Function

AggiornaDati_ed_EsportaDati_Pfunct_Err:
    MsgBox Error$
    Resume AggiornaDati_ed_EsportaDati_Pfunct_Exit

End Function







'// *** FINE ***
'//             AGGIORNA I DAITI CON LE QUERY ED ESPORTA IN MDB + XLS
'//****************************************************************************************************************//




'// *** FINE ***
'//             I COMANDI CMD :
'//                             DoCmdOpenQuery_Sub
'//
'//****************************************************************************************************************//
'//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@//

'//DOCMD_OPEN_QUERY = Eseguo il comando docmd di esecuzione della query la cui _
                    esistenza viene controllata prima di chiamare questa routine. _
                    I parametri sono due: _
1 paramentro par_acViewNormal, corrisponde alla variabile AcView sono i seguenti e devono essere indicati in par_acViewNormal _
Nome Valore         Descrizione _
acViewDesign        1 Visualizzazione Struttura _
acViewLayout        6 Visualizzazione Layout _
acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
acViewPivotChart    4 Visualizzazione Grafico pivot _
acViewPivotTable    3 Visualizzazione Tabella pivot _
acViewPreview       2 Anteprima di stampa _
acViewReport        5 Visualizzazione Report _

'Il 2 parametro, par_acEdit Deve indicare le modalita di apertura dell'oggetto: _
Costante di enumerazione AcOpenDataMode _
Specifica la modalità di immissione dati all'apertura di una tabella, una query, _
una stored procedure, una visualizzazione o una funzione definita dall'utente. _
Nome Valore     Descrizione _
acAdd               0       L'utente può aggiungere nuovi record ma non visualizzare né modificare i record esistenti. _
acEdit              1       L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record. _
acReadOnly          2       L'utente può solo visualizzare i record.



Public Sub DOCMDOpenQuery_Sub(par_Query_s As String, par_acViewNormal As Integer, par_acEdit As Integer)

                    '//Eseguo la query passata come parametro solo se valorizzata _
                       la variabile che contiene il nome della query. Il controllo se _
                       la query esiste viene fatto a prima della routine.
                    If par_Query_s > "" Then
                            DoCmd.OpenQuery par_Query_s
                    Else
                        MsgBox "Comando Cmd.OpenQuery non eseguito in quanto il parametro oggetto 'par_Query_s' è nullo:  --> " & par_Query_s
                    End If
                    
End Sub




















                                    '//*************************************//
                                    '//     LE ESPORTAZIONI
                                    '//*************************************//
                                    '
'//**************************************************************************************************************************//
'//         FUNZIONI DI ESPORTAZIONI LAVORI PUBBLICI IN FORMATO XLS
'//**************************************************************************************************************************//
'//

'//01)_ESPORTA FILE XLS CON MATRICE = Funzione di esportazione dei dati dei lavori pubblici _
        con esportazioni in tabelle excel. Viene utilizzata una matrice.
'//     CODICE------> ESPORTA_FILE_XLS_ConMatriceDati_pFunct



'//ESPORTO SU FILE ESTERNO _
per ora un xls
Public Sub OutputTo_Sub(par_Tabella_s As String, par_Path_s As String, par_NuovaDenominazioneTabella_01_s As String, par_Bool As Boolean)
  
        '//MODELLO DI ESPORTAZIONE
        '//Esporto le tabelle in formato xls
        '                       par_Tabella_s                          par_PATH_S                               par_Bool
        'DoCmd.OutputTo acTable, Tabella_01_s, "MicrosoftExcel(*.xls)", PATH_s & NuovaDenominazioneTabella_01_s, False, ""
        'MODELLO FUNZIONANTE ---> DoCmd.OutputTo acTable, "LLPP_ATTI_Tb01_Gestione", "MicrosoftExcel(*.xls)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LPP_ARCHIVI_SALVATAGGI\LLPP_ATTI_Tb01_Gestione.xls", False, ""
        '//Controllo parametri:
            Debug.Print "par_Tabella_s                              ----> " & par_Tabella_s
            Debug.Print "par_PATH_S                                 ----> " & par_Path_s
            Debug.Print "par_NuovaDenominazioneTabella_01_s         ----> " & par_NuovaDenominazioneTabella_01_s
            
        DoCmd.OutputTo acTable, par_Tabella_s, "MicrosoftExcel(*.xls)", par_Path_s & par_NuovaDenominazioneTabella_01_s, par_Bool, ""

End Sub






'//**************************************************************************************************************************//

'//FUNZIONE ESPORTA FILE XLS
'//Chiamo matrice con Stringhe che rappresentano delle tabelle _
Impegni , Presenze
Private Sub CHIAMA_ESPORTA_FILE_XLS_ConMatriceDati_pFunct()
    
    '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls
    ESPORTA_FILE_XLS_ConMatriceDati_pFunct "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA", _
    "PRES3000_Tb01_Calendario", "PRES3000_Tb02_ElencoGiornaliere", "PRES3000_Tb01_Calendario", "Indirizzi_Tb01_INTESTATARI", _
    "Indirizzi_Tb02_ELENCO"

End Sub


'//01)_ESPORTA FILE XLS CON MATRICE = Funzione di esportazione dei dati dei lavori pubblici _
        con esportazioni in tabelle excel. Viene utilizzata una matrice.
'//     CODICE------> ESPORTA_FILE_XLS_ConMatriceDati_pFunct
'//=================================================================================================//
'//NOTE ------------->  : attenzione esporto la collezione di file ma il primo deve essere nullo ("") per _
                        evitare che venga saltato.

Public Function ESPORTA_FILE_XLS_ConMatriceDati_pFunct(strName As String, ParamArray intScores() As Variant)
Dim Path_s As String
Dim Tabella_01_s As String
Dim NuovaDenominazioneTabella_01_s As String
Dim Tabella_02_s As String
Dim NuovaDenominazioneTabella_02_s As String
Dim intI As Integer
Dim iCount As Integer                   '//il contatore

On Error GoTo ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Err


'//ESPORTA I DATI IN FILE XLS
'//........................................................
 


   '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls _
      ATTENZIONE, occorre inserire un record vuoto "", affinche possono essere considerati tutti gli altri _
      file dell'esportazione altrimenti di viene saltato il primo "LLPP_ATTI_Tb01_Gestione", _
                Call ESPORTA_FILE_XLS_ConMatriceDati_pFunct "", _
                "LLPP_ATTI_Tb01_Gestione", _
                "LLPP_ATTI_Tb02_Allegati", _
                "Indirizzi_Tb01_INTESTATARI", _
                "Indirizzi_Tb02_ELENCO", _
                "PRES3000_Tb01_Calendario", _
                "PRES3000_Tb02_ElencoGiornaliere", _
                "LLPP_DF01_CodiceOpera", _
                "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", _
                "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"

'//........................................................


'//Reset ed impostazioni
'//........................................................
    Path_s = vbNull
    Path_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LPP_ARCHIVI_SALVATAGGI\"
    
    Tabella_01_s = vbNull
    Tabella_01_s = "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO"

'//........................................................

    Debug.Print
    Debug.Print strName; " Punteggi"
    ' Utilizza la funzione UBound per definire il
    ' limite superiore della matrice.

    '//Dimensione la matrice di file
    
     Dim MatriceFile_s As String
    
    iCount = 0
    
    
   
            '//ATTIVAZIONE CLASSE
            '//ROUTINE_DI_ATTIVAZIONE_APERTURA_FILE
            '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva.LaClasse
            '//---------------------------------------------------------------------------------//
            '//NOTA  :Controlla se esiste la form negli oggetti del database
                
            
                    'If Str1 <> "OGGETTO NON TROVATO" Then
                         '
                            'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As New CLASSE_ATTI_Tb01_Gestione
                            'Dim OggettoDaRicercare_s As String
                            
                            'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione
                        
                        
              
                    'End If
                                                    
                '//---------------------------------------------------------------------------------//
    
    
    
    
    
    
    
    For intI = 0 To UBound(intScores())
        Debug.Print "          "; intScores(intI)
        Tabella_01_s = vbNull
        Tabella_01_s = intScores(intI)
        iCount = iCount + 1
        Bool1 = False
        NuovaDenominazioneTabella_01_s = intScores(intI) & ".xls"
        
               
            '//Controllo Parametri e tabella da esportare
            Debug.Print
            Debug.Print "----------------------CONTROLLO VARIABILI TABELLE XLS --------------------------------------"
            Debug.Print
            Debug.Print "tabella da ricercare - variabile           :Tabella_01_s                                   ----> " & Tabella_01_s
            Debug.Print "la path e la variabile                     :par_PATH_S                                     ----> " & Path_s
            Debug.Print "nuova denominazione xls con la variabile   :NuovaDenominazioneTabella_01_s                 ----> " & NuovaDenominazioneTabella_01_s
            Debug.Print
            Debug.Print "--------------------------------------------------------------------------------------------"
               
               
                '//ESPORTO LA TABELLA
                '//.......................................................................................
                '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando. _
                        la TABELLA viene presa nella matrice.
                    
                    '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                       di ricerca secondo il tipo di oggetto
                       
                    Bool1 = False
                    'Tabella da ricercare = viene indicata nella matrice intScores(intI) senza .xls
                    Bool1 = MyCls_CLASSE_ATTI_Tb01_Gestione.Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct(Tabella_01_s)
                    
                        '//SE TRUE = oggetto esistente ed eseguo il comando
                        If Bool1 Then
                            
                            '???? NON FUNZIONA
                            'DoCmd.OutputTo acTable, Tabella_01_s, "MicrosoftExcel(*.xls)", PATH_s, NuovaDenominazioneTabella_01_s, False, ""
                            
                            
                            '//ESEGUO IL COMANDO OutpuTo CON con 3 parametri, il nome della TABELLA + LA PATH dove esportare + TRUE ?
                            'MODELLO ---> OutputTo_Sub(par_Tabella_s As String, par_PATH_S As String, par_NuovaDenominazioneTabella_01_s As String, par_Bool As Boolean)
                            MyCls_CLASSE_ATTI_Tb01_Gestione.OutputTo_Sub Tabella_01_s, Path_s, NuovaDenominazioneTabella_01_s, False
                            
                            '//APRO TABELLA ESPORTATA
                            DoCmd.OpenTable Tabella_01_s, acViewNormal, acEdit
                                 
                                                
                        
                      End If
                        
               '//.......................................................................................

        
                
                    MatriceFile_s = MatriceFile_s & iCount & ") " & NuovaDenominazioneTabella_01_s & " " & vbCrLf
        
        '//Esporto le tabelle in formato xls
        '                       par_Tabella_s                          par_PATH_S                               par_Bool
        'sospesa DoCmd.OutputTo acTable, Tabella_01_s, "MicrosoftExcel(*.xls)", PATH_s, NuovaDenominazioneTabella_01_s, False, ""


    Next intI
    
    '//RILASCIO L'OGGETTO CLASSE
    Set CLASSE_ATTI_Tb01_Gestione = Nothing
                            
                        
    '//Messaggio di fine attivita'
    MsgBox "ESPORTAZIONE DATI IN FILE XLS EFFETTUATI in --> " & Path_s & vbCrLf & MatriceFile_s & vbCrLf
    

ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Exit:
    Exit Function

ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Err:
    MsgBox Error$
    Resume ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Exit

End Function
'//01)_ESPORTA FILE XLS CON MATRICE = Funzione di esportazione dei dati dei lavori pubblici   ***FINE***
'//=================================================================================================//

'//ESPORTA_FILE_XLS_ConMatriceDati_pFunct
'//*************************************************************************************************//
'//NOTE : ESPORTA i file in formato xls (csv)   *** FINE ***
'//*************************************************************************************************//




                                            '//*** FINE ***
                                    '//*************************************//
                                    '//     LE ESPORTAZIONI
                                    '//*************************************//

'//**************************************************************************************************************************//
'//         FUNZIONI DI ESPORTAZIONI LAVORI PUBBLICI IN FORMATO XLS
'//**************************************************************************************************************************//



                                    '//***INIZIO***
                                    '//*************************************//
                                    '//     LE FUNZIONI GENERALI
                                    '//*************************************//
'//********************************************************************************************************************************************************//
'//                                 LE FUNZIONI DI GESTIONE ATTI Tb01_Gestione
'//********************************************************************************************************************************************************//
'//
'//NOTA: In questa classe vengon raggruppate tutte le fuzioni che attengono alla gestione della TABELLA ATTI _
        per la gestione dei lavori pubblici. _
        IL seguente elenco specifica le funzioni implementate:
        
        
'//01.AGGIORNA_FILE_ATTO_pfunct = Aggiorna file atto crea il NOME DI FILE da salvare composto da _
                               fileAtto+NroAtto+Annoatto, solo se nro atto e data sono > 0 _
                               CODICE------>AGGIORNA_FILE_ATTO_pfunct.attiva
        
'//********************************************************************************************************************************************************//
'//********************************************************************************************************************************************************//


    
    '//01.AGGIORNA_FILE_ATTO_pfunct = Aggiorna file atto crea il NOME DI FILE da salvare composto da _
                                    fileAtto+NroAtto+Annoatto, solo se nro atto e data sono > 0 _
                                    CODICE------>AGGIORNA_FILE_ATTO_pfunct.attiva
    '//===========================================================:===================================:================================
    '//AGGIORNA_ATTI_Mdl_N001_001_Function.001.01__________:Funzione 00 CALCOLA                                :Calcola
    '//...........................................:.....................................................................................
    '//NOTA PROCEDURA: La Funzione Aggiona il campo FILE_ATTO ricostruendo il nome del file con i parametri passati.
    
    Public Function AGGIORNA_FILE_ATTO_pfunct(par_TipoAtto_s As String, _
                                              par_NroAtto_lng As Long, _
                                              par_DataAtto_d As Date, _
                                              par_AnnoAtto_i As Integer) As String
    
    '//.......................................................
    '// DIM VARIABILI
    'Dim sScelta_db As String
    
        
        
        
        '//MODELLO-> AGGIORNA_FILE_ATTO
        '//--------------------------------------------------------------------------//
        '//NOTA   : Aggiorna il file atto costruendo il nome del file _
                   con i parametri passati ottengo il nome dai seguenti dati _
                   FILE_ATTO=TipoAtto+NroAtto+Annoatto a condizione che _
                   il NroAtto >0 and DataAtto >"", altrimenti la funzione restituisce _
                   il valore nullo. Vengono controllati anche le variabili Nro atto, _
                   data atto ed anno anno che devono essere sempre valorizzate.
                   
                   '//RESETTO IL FILE ATTO _
                   in questo modo posso riattivare l'evento anche nel caso di errori nella _
                   digitazione del numero file atto, potendo ricostruire in automatico il _
                   nome completo del File_atto
                   'Me.FileAtto_s_Txt = ""
                    
                    '//CONTROLLO VARIABILE E CHIAMATA PROCEDURA
                    '//.........................................................//
                        'sSTR1 = ""
                        'TipoAtto_s = Me.TipoAtto_s
                        'If Me.NroAtto_lng > 0 Then
                        '    NroAtto_lng = Me.NroAtto_lng
                        'Else
                        '    Me.NroAtto_lng = 0
                        'End If
                        
                        '//CONTROLLO DATA = se la data nulla viene recuperato quella di calendario.
                        'If IsNull(DataAtto_d) = TRUE Then
                        '    DataAtto_d = Date
                        'Else
                            '//Altrimenti viene recuperata quella della casella di testo valorizzata.
                        '    DataAtto_d = Me.DataAtto_d
                        'End If
                        
                        '//Controllo variabile Anno deve essere valorizzata
                        'If AnnoAtto_i = 0 Or IsNull(AnnoAtto_i) = True Then
                            '//Se NULL o Zero, viena applicato l'anno corrente
                        '    AnnoAtto_i = Year(Date)
                        'Else
                            '//Altrimenti viene recuperato l'anno del campo testo
                        '    AnnoAtto_i = Me.AnnoAtto_i
                        'End If
                        
                        
                        
                        '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                        'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione
                        
                        '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
                        l'oggetto di classe alla variabile e creo creo l'oggetto di classe
                        'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione
                        
                         '//MODELLO FUNZIONANTE
                         'sSTR1 = AGGIORNA_FILE_ATTO_pfunct(TipoAtto_s, NroAtto_lng, DataAtto_d, AnnoAtto_i)
                         
                        
                        'sSTR1 = MyCls_CLASSE_ATTI_Tb01_Gestione.AGGIORNA_FILE_ATTO_pfunct(TipoAtto_s, NroAtto_lng, DataAtto_d, AnnoAtto_i)
                        
                        '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                        'Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing

                
                        'Me.DATAAGG_TXT = Date
                        'Me.TIMEOPER_TXT = Time()
                        
                    '//.........................................................//
                
        '//MODELLO-> AGGIORNA_FILE_ATTO *** FINE ***
        '//--------------------------------------------------------------------------//
          
        
        
        
            On Error GoTo Err_AGGIORNA_FILE_ATTO_pfunct
        
        
            '//AGGIORNA_FILE_ATTO
            '//-------------------------------------------:-----------------------------------------------
            '//AGGIORNA_ATTI_Mdl_N001_001_Function.001.02__________:
            '//NOTA   : Aggiorna il file atto costruendo il nome del file con i parametri passati.
            
                                                     
                            '//AGGIORNAMENTO CAMPO NRO ATTO -
                            '//..............................................................
                            '//Aggiorna il nro att and data > 0 allora _
                            creo il nome del file con fileAtto+NroAtto+Annoatto, solo se _
                            nro atto e data sono > 0
                            
                                If par_NroAtto_lng > 0 And par_DataAtto_d > 0 Then
                                    Str1 = ""
                                    Str1 = par_TipoAtto_s & "_" & par_NroAtto_lng & "_" & par_AnnoAtto_i
                                    '//controllo e salvataggio
                                    Debug.Print "-------------------------------------------------------"
                                    Debug.Print "file creato"
                                    Debug.Print Str1
                                    Debug.Print "-------------------------------------------------------"
                                    AGGIORNA_FILE_ATTO_pfunct = Str1
                                    
                                    
                                End If
                                
                            '//..............................................................
                    
            '//-------------------------------------------:-----------------------------------------------
            
    
    
    
    
    
    
    '//--------------------------------------------------------------------------------------------------
    '//                       FINE FUNCTION E GESTIONE ERRORI
    
Exit_AGGIORNA_FILE_ATTO_pfunct:
    Exit Function
    
Err_AGGIORNA_FILE_ATTO_pfunct:
        MsgBox "ERRORE FUNCTION PUBLIC    " & Err.number & " - " & Err.Description, vbCritical, "AGGIORNA_FILE_ATTO_pfunct"
        Resume Exit_AGGIORNA_FILE_ATTO_pfunct
     
    End Function
    '//AGGIORNA_FILE_ATTO_pfunct                                              *** FINE ***
    '//===========================================================:===================================:================================
    
    
    



'//********************************************************************************************************************************************************//
'//                                 LE FUNZIONI DI GESTIONE ATTI Tb01_Gestione
'//********************************************************************************************************************************************************//

                                    '//***FINE***
                                    '//*************************************//
                                    '//     LE FUNZIONI GENERALI
                                    '//*************************************//











