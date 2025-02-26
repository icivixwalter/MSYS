'//MODULO
'//*******************************************************************************************************//
'// GESTIONE_Mdl10_FILE_APRI_PDF
'//*******************************************************************************************************//
'//NOTE: MODULO STANDART DI GESTIONE DELLE API che deve essere seprato dalla form GEFILE_Frm01_GESTIONI
'//      quindi il modulo gestisce le API WINDOWS mentre la form GEFILE... lo richiama. _
         @codice---> @GESTIONE@FILE@API@WINDOWS


Option Compare Database

Dim Lng As Long
Dim Path_programma_s As String
Dim NameFile_s As String




Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000


'//-----------------------------------------------------------------------------------//
'//NOTE:  Precompilazione della routine a 64 e 32 bit
#If VBA7 Then
    ' Windows a 64 bit
    Private Declare PtrSafe Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
        "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath As String, _
        ByVal lpOutPath As String) As Long
#Else
    ' Windows a 32 bit
    Private Declare Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
        "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath As String, _
        ByVal lpOutPath As String) As Long
#End If


#If VBA7 Then
    ' Windows a 64 bit
    Public Declare PtrSafe Function GetOpenFileNameA Lib "comdlg32.dll" (ByRef lpofn As OPENFILENAME) As Long
#Else
    ' Windows a 32 bit
    Public Declare Function GetOpenFileNameA Lib "comdlg32.dll" (ByRef lpofn As OPENFILENAME) As Long
#End If






' Dichiarazione della struttura OPENFILENAME
Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long            ' Usa Long
    hInstance As Long            ' Usa Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type







'//routine del modulo per l'aperutura dei pdf: @APRI_PDF_MODULO->GEFILE_Mdl01_GESTIONI_FILE_API_VBA
'//ATTIVO LA SUB APRI SINGOLO FILE PDF
Public Sub APRI_FILE_PDF_pSub()

    '//La Path del programma per esteso
    Path_programma_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\PDF\"
    '//Il nome del file in pdf
    NameFile_s = "FILTRO_SQL.pdf"
    
    
    '//chiamo la sub con i parametri
    ApriFilePdf Path_programma_s, NameFile_s, 100
    
End Sub

'//@MODELLO@APRI@FILE_(questo è il modello per l'apertura dei file e funziona su tutto con 3 parametri)
'//==================================================================================================================//
'//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
'//TIPO------------------> ROUTINE pubblica.
'//--------------------------------------------------------------@
'//PARAMETRI            -> passo 2 stringhe per parametro, la path e il nome del file.pdf + ID GESTIONE
'//VALORE_DI_RITORNO    -> Nulla
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.00



Public Function ApriFilePdf(par_Path_s As String, par_NameFile_s As String, par_IDGestione_lng As Long) As String

Dim Path_programma_s As String
Dim NameFile_s As String
Dim Stringa1 As String

On Error GoTo ApriFilePdf_Err
        

    
                        
            '//MODELLO DI APERTURA DEL FILE SPECIFICO CON IL COMANDO OGGETTO PDF - @MODELLO@APRI.FILE
            '//--------------------------------------------------------------------------------//--------//
            '//NOTE                 -> Apro il file di tipo doc, zip o pdf utilizzando la procedura pubblica _
                                     del MODULO : GESTIONE_Mdl10_FILE_APRI_PDF
            '//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.01
            '//PARAMETRI            -> par_Path_s                   = PATH _
                                    -> par_NameFile_s               = NOME FILE _
                                    -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                                    
                                    
              '//IMPOSTO LE VARIABILI PATH E FILE
              '//Dim MyFile_s As String
              '//Dim MyPath_s  As String
              
                    '//imposto i parametri - LA PATH
                    'MyPath_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\"
                    '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
                    'MyFile_s = "Project_PROGETTI_MDB.sublime-project "
                        
                        
                     '//chiamo la sub con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng)
                        'Call ApriFilePdf(MyPath_s, MyFile_s, 0)
                         
                    
            '//--------------------------------------------------------------------------------//--------//
                                    
                                    


   '//APRO FILE PDF SPECIFICO
   '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//
   '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
   '//           incluso il file PDF.
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
                        GoTo ApriFilePdf_Exit
                    End If
    
                    
                    
    '
              
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
      
      
ApriFilePdf_Exit:
    Exit Function

ApriFilePdf_Err:
    MsgBox Error$ & " ERRORE APRI FILE : " & ProceduraMessaggioErrore_s & " " & ProceduraAttivaEseguita_s
    Debug.Print ProceduraMessaggioErrore_s & " " & ProceduraAttivaEseguita_s

    Resume ApriFilePdf_Exit

End Function

'//APRI_IL_FILE_PDF         *** FINE ***
'//==================================================================================================================//







'//@API@FUNZIONE@CERCA@FILE_(le istruzioni api per la ricerca dei FILE)
'//===============================================================================================//
'//NOTE:        funzione di ricerca dei file che utilizza
'//CODICE --->  @apiSearchTreeForFile.PRECOMPILA

'//CODICE ---> @01@PRECOMPILAZINE@FUNZIONE@API

'//ATTIVO LA FUNZIONE DI PROGA= non simmmmoooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
'//-----------------------------------------------------------------------------------//
'//'//CODICE --->  @02@apiSearchTreeForFile.chiama@sub
Private Sub CHIAMA_fReturnFilePath()
    fSearchFile "LLPP_IMPEGNI_ATTIVA_GESTIONE.mdb", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\IMPEGNI_DI_SPESA\MDB\"
End Sub
'//-----------------------------------------------------------------------------------//

'//FUNZIONE CERCA FILE con le api windows
'//-----------------------------------------------------------------------------------//
'//
'//CODICE ---> @03@apiSearchTreeForFile.FUNZIONE

Function fSearchFile(ByVal strFilename As String, ByVal strSearchPath As String) As String
    Dim lpBuffer As String
    Dim lngResult As Long

    strFilename = strFilename & "*.*"
    fSearchFile = ""
    lpBuffer = String$(1024, 0)
    
    '//CODICE ---> @03@apiSearchTreeForFile.FUNZIONE.01@chiave.funzione
    lngResult = apiSearchTreeForFile(strSearchPath, strFilename, lpBuffer)
    If lngResult <> 0 Then
        
        '//CODICE ---> @03@apiSearchTreeForFile.FUNZIONE.02@ricerca.file
        If InStr(lpBuffer, vbNullChar) > 0 Then
            fSearchFile = Left$(lpBuffer, InStr(lpBuffer, vbNullChar) - 1)
            Debug.Print "---------------------------------------------------"
            Debug.Print "Controllo file "
            Debug.Print fSearchFile
            Debug.Print "---------------------------------------------------"
            
            MsgBox fSearchFile & Chr$(13) _
               & "01 Path : -> " & strSearchPath & Chr$(13) _
               & "02 File : -> " & strFilename & Chr$(13) _
               & " Risultato -> FILE TROVATO!"
        End If
    Else
        MsgBox fSearchFile & Chr$(13) _
               & "  02) Path : -> " & strSearchPath & Chr$(13) _
               & "  02) File : -> " & strFilename & Chr$(13) _
               & " Risultato -> FILE NON TROVATO!"
    End If
End Function
'//-----------------------------------------------------------------------------------//

'//===============================================================================================//





