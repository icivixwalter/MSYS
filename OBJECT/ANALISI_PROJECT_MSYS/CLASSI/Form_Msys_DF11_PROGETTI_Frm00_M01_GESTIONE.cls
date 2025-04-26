VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Msys_DF11_PROGETTI_Frm00_M01_GESTIONE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim ORDINA_B As Boolean         '//VARIABILE GENERALE altrimenti non funziona.


Private Sub Form_Open(Cancel As Integer)
    
Me.Recalc

End Sub

Private Sub IdOrd_lng_TXT_Click()

        '//ordinamento

End Sub

Private Sub IdOrd_lng_TXT_DblClick(Cancel As Integer)


    '//ORDINAMENTO FALSE = DISCENDENTE
    '//--------------------------------------------------------------------------------//--//
    If ORDINA_B = False Then
        
        Me.Form.OrderBy = "[IdOrd_lng] DESC"
        Me.Form.OrderByOn = True
        ORDINA_B = True
        
        '//ORDINAMENTO TRUE = ASCENDENTE
    Else
        ORDINA_B = True
        Me.Form.OrderBy = "[IdOrd_lng] ASC"
        Me.Form.OrderByOn = True
    End If
             Me.Recalc
    '//--------------------------------------------------------------------------------//--//
 
End Sub


Private Sub Cmd_Open_Click()

On Error GoTo Err_Cmd_Open_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    
    
        '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
        '//--------------------------------------------------------------------------------//--------//
        '//NOTE                 -> Apro il file di tipo doc, zip o pdf
        '//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.01
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
                MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\HELP\"
                '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
                '
                MyFile_s = "HELP_ANALISI_ESPORTA_MSYS.odt"
                    
                    
                 '//chiamo la sub con i parametri =   'CALL (ApriFilePdf_Interna_s_PFunct Path_programma_s, NameFile_s, IDGestione_lng)
                    '
                    Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
                     
                
        '//--------------------------------------------------------------------------------//--------//
  

Exit_Cmd_Open_Click:
    Exit Sub

Err_Cmd_Open_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_Open_Click
    
End Sub




'//@MODELLO@APRI@FILE_(questo è il modello per l'apertura dei file e funziona su tutto con 3 parametri)
'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE -
'//==================================================================================================================//
'//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
'//TIPO------------------> ROUTINE pubblica.
'//--------------------------------------------------------------@
'//PARAMETRI            -> passo 2 stringhe per parametro, la path e il nome del file.pdf + ID GESTIONE
'//VALORE_DI_RITORNO    -> Nulla
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.00


Public Function ApriFilePdf_Interna_s_PFunct(par_Path_s As String, par_NameFile_s As String, par_IDGestione_lng As Long) As String

Dim Path_programma_s As String
Dim NameFile_s As String
Dim Stringa1 As String

On Error GoTo ApriFilePdf_Interna_s_PFunct_Err


   
'//MODELLO DI APERTURA DEL FILE SPECIFICO CON IL COMANDO OGGETTO PDF - @MODELLO@APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//Dim MyFile_s As String
  '//Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        'MyPath_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\"
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        'MyFile_s = "Project_PROGETTI_MDB.sublime-project "
            
            
         '//chiamo la sub con i parametri =   'CALL (ApriFilePdf_Interna_s_PFunct Path_programma_s, NameFile_s, IDGestione_lng)
            'Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
        
'//--------------------------------------------------------------------------------//--------//
                        
    
    
    
    '//APRO_FILE
    '//--------------------------------------------------------------------------------//-------//
    '//APRO_FILE_O_PATH = Attivo la routine di apertura del file o di una path. La procedura _
            utilizza un eseguibile java.exe a cui viene passato un parametro composto di due _
            elementi il nome dell'eseguibile java + NOME FILE con la path estesa  e la PATH _
            che localizza dove si trova l'eseguibile java.exe
    
            '//CODICE -----------> APRO_FILE_O_PATH_N01_pSub.01.02
    
                '//RESET
                ProceduraMessaggioErrore_s = ""
                ProceduraAttivaEseguita_s = ""
   

                '//IMPOSTO LE VARIABILI
                ProceduraMessaggioErrore_s = "ERRORE NELLA ROUTINE Per questa path e file: " & par_Path_s & par_NameFile_s
                ProceduraAttivaEseguita_s = "ApriFilePdf_Interna_s_PFunct"

    
               '//controllo parametri
                    If par_Path_s = "" Then
                        MsgBox "PARAMETRI APERTURA FILE VUOTI uscita dalla Routine", vbExclamation
                        GoTo ApriFilePdf_Interna_s_PFunct_Exit
                    End If
    
                    
                    
    '//--------------------------------------------------------------------------------//-------//
    
                        
         
                        '//APRO FILE PDF SPECIFICO
                        '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//
                        '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
                        '//           incluso il file PDF.
                          
                          '//La path del file = Unisco la path + il file+estensione sopra individuata
                          par_Path_s = par_Path_s & par_NameFile_s & MyFileEstensione_s
                                 
                          '//controllo
                          Debug.Print par_Path_s
                                 
                          '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                          '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                          '// prima la sessione dos e poi il pdf..
                          '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                          Set WshPDF = CreateObject("wscript.shell")
                          
                          '// Unisco il comando pdf start + path definitiva
                          Stringa1 = "%comspec% /c start " & par_Path_s                               '//Comando di apertura
                          '//Attivo il comando
                          WshPDF.Run Stringa1
                                  
                                  
                          '//libero la memoria dalle variabili e dagli oggetti creati
                          WshPDF = Null
                          Stringa1 = vbNull
                          MyFileEstensione_s = vbNull
                          MyPath_s = vbNull
                            
                            '//sistemare ??? invio tasti ???
                            Lng = 0
                            For Lng = 1 To 90000000 Step 1
                            
                            Next Lng
                            
                          'SendKeys "{ESC}"
                          'SendKeys "{ESC}"
                          'SendKeys "{ESC}"
                          
                            
                        '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//

       '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
            '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
            '// prima la sessione dos e poi il pdf..
            '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
            'Set WshPDF = CreateObject("wscript.shell")
            
            '// Unisco il comando pdf start + path definitiva
            'Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
                '//Attivo il comando ed aggiungo @exit per @chiudere@il@terminale occorre lo spazio per evitare l'errore & " ^exit"
            'WshPDF.Run Stringa1 & " ^exit"
      
      
ApriFilePdf_Interna_s_PFunct_Exit:
    Exit Function

ApriFilePdf_Interna_s_PFunct_Err:
    MsgBox Error$
    
    Debug.Print ProceduraMessaggioErrore_s & " " & ProceduraAttivaEseguita_s

    Resume ApriFilePdf_Interna_s_PFunct_Exit

End Function

'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE -         *** FINE ***
'//==================================================================================================================//






'//DENOMINAZIONE---------> APRO_FILE_O_PATH_N01_pSub
'//=================================================================================================================//
'//Parametri che possono essere cambiati collettivamente:
'//--------------------------------------------------------------@
'//TIPO------------------> ROUTINE pubblica.
'//ATTIVITA--------------> APRO_FILE_O_PATH
'//ProceduraMessaggioErrore_s = "APRO_FILE_O_PATH"
'//ProceduraAttivaEseguita_s = "APRO_FILE_O_PATH_N01_pSub"
'//...........................
'//--------------------------------------------------------------@
'//NOTE------------------> ....
'//PARAMETRI-------------> Nessuno
'//VALORE_DI_RITORNO-----> Nulla
'//CODICE----------------> APRO_FILE_O_PATH_N01_pSub.01.EseguoRoutine
'//=================================================================================================================//
Private Sub APRO_FILE_O_PATH_N01_pSub(par_Path_s As String, par_File_s As String)

On Error GoTo APRO_FILE_O_PATH_N01_pSub_Err

      
    
    '//APRO_FILE_O_PATH = Attivo la routine di apertura del file o di una path. La procedura _
            utilizza un eseguibile java.exe a cui viene passato un parametro composto di due _
            elementi il nome dell'eseguibile java + NOME FILE con la path estesa  e la PATH _
            che localizza dove si trova l'eseguibile java.exe
    '//--------------------------------------------------------------------------------//-------//
    '//CODICE -----------> APRO_FILE_O_PATH_N01_pSub.01.02
    
                '//RESET
                ProceduraMessaggioErrore_s = ""
                ProceduraAttivaEseguita_s = ""
   

                '//IMPOSTO LE VARIABILI
                ProceduraMessaggioErrore_s = "ERRORE NELLA ROUTINE APRO_FILE_O_PATH"
                ProceduraAttivaEseguita_s = "APRO_FILE_O_PATH_N01_pSub"

    
               '//controllo parametri
                    If par_Path_s = "" Then
                        MsgBox "PARAMETRI APERTURA FILE VUOTI uscita dalla Routine", vbExclamation
                        GoTo APRO_FILE_O_PATH_N01_pSub_Err_Exit
                    End If
    
                    PARAMETRO_JAVA_EXE_s = par_Path_s & par_File_s
                    
    '//--------------------------------------------------------------------------------//-------//
    
    
    
                    
                    
                              '//LA PATH DELL'ESSEGUIBILE JAVA E' UTILIZZATA PER TUTTE LE CHIAMATE
                              '//...................................................................................//
                                  '//ESEGUIBILE = il file java.exe che viene ATTIVA_APRI_FileOPath_Subto _
                                     Myfile_s è costituito da due elementi, l'esguibile java.exe e il parametro inviato _
                                     all'eseguibile stesso. Per richiamare l'eseguibile occorre passare la path dove _
                                     si trova il java.exe in MyPath_s.
                                  MyFile_s = "Esegui_Programma_(ConParametro).exe" & " " & PARAMETRO_JAVA_EXE_s
    
                                  '//LA PATH = dove si trova l'eseguibile (cartella ZZ_ESEGUI_EXE)
                                  MyPath_s = "c:\Casa\CDM\GE_CDM\ZZ_ESEGUI_EXE\"
                                         
                                            
                                  'MyFileJavaExe_s = "Esegui_Programma_(ConParametro).exe" & par_PARAMETRO_JAVA_EXE_s
    
                                  '//MyPathJavaExe_s = dove si trova l'eseguibile java.exe da ATTIVA_APRI_FileOPath_Subre(nella cartella ZZ_ESEGUI_EXE)
                                  'MyPathJavaExe_s = "c:\Casa\CDM\GE_CDM\ZZ_ESEGUI_EXE\"
                                            
                                            Debug.Print "       CONTROLLO PARAMETRI JAVA.EXE                        "
                                            Debug.Print "-----------------------------------------------------------"
                                            Debug.Print "Eseguibile java.exe + LA PATH COMPLETA con il NOME DEL FILE in MyFileJavaExe_s    : "
                                            Debug.Print MyFileJavaExe_s
                                            Debug.Print
                                            Debug.Print "LA PATH = dove si trova l'eseguibile (cartella ZZ_ESEGUI_EXE)- MyPathJavaExe_s    : "
                                            Debug.Print MyPathJavaExe_s
                                            Debug.Print
                                            Debug.Print par_PARAMETRO_JAVA_EXE_s
                                            
                                            Debug.Print "-----------------------------------------------------------"
                                        
                                
                              '//...................................................................................//
                              
                              
                              
                              
    
        
        '//ATTIVO LA CLASSE PER L'APERTURA DELLA PATH O FILE GENERICI
        '//.....................................................................................................
        '//Passo un parametro stringa che può rappresentare una path da aprire esempio : _
           "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO". _
           Se voglio aprire un file  occore passare la PATH+FILE con estensione es. _
           "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO\file.xls"
                   
            '//File da ricercare e la path
            'Dim MyFileJavaExe_s As String
            'Dim MyPathJavaExe_s As String
            
            'MyFileJavaExe_s = "LLPP_ATTI_ALLEGATI_ARCHIVIO.RAR"
            'MyPathJavaExe_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\LPP_ALLEGATI\"
           
        
             '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                '
                    Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione
        
            '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
               l'oggetto di classe alla variabile e creo creo l'oggetto di classe
                '
                    Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione
                    
                    
                    Debug.Print "controllo parametri"
                    Debug.Print MyPathJavaExe_s
                    Debug.Print MyFileJavaExe_s
            
            '//PATH E FILE = Inserire LA PATH ED IL FILE con percoso esteso.
                    'MyCls_CLASSE_ATTI_Tb01_Gestione.APRO_PATH_O_FILE_GENERICI_pSub (MyPathJavaExe_s & MyFileJavaExe_s)
                    MyCls_CLASSE_ATTI_Tb01_Gestione.APRO_PATH_O_FILE_GENERICI_pSub (MyPath_s & MyFile_s)
                
                
            '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                    Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing
        
        '//.....................................................................................................
       
                              
                              
                              
             
    
    
    
    
    
APRO_FILE_O_PATH_N01_pSub_Err_Exit:
        Debug.Print "uscita dalla routine : " & ProceduraMessaggioErrore_s
        Debug.Print ProceduraAttivaEseguita_s
   
    Exit Sub

APRO_FILE_O_PATH_N01_pSub_Err:
    MsgBox Error$
    Resume APRO_FILE_O_PATH_N01_pSub_Err_Exit
    MsgBox Err.Description & " " & ProceduraMessaggioErrore_s & " - " & ProceduraAttivaEseguita_s & " --> errore Error$ -> : " & Error$
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Debug.Assert "BLOCCO PROCEDURA -> " & ProceduraAttivaEseguita_s
  
    Stop
    Resume APRO_FILE_O_PATH_N01_pSub_Err
    
End Sub
'//'//DENOMINAZIONE---------> APRO_FILE_O_PATH_N01_pSub *** FINE ***
'//=================================================================================================================//

