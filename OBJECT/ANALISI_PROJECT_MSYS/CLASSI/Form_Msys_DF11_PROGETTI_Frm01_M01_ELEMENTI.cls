VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Msys_DF11_PROGETTI_Frm01_M01_ELEMENTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database




Private Sub HELP_ANALISI_ESPORTA_MSYS_Click()

On Error GoTo Err_HELP_ANALISI_ESPORTA_MSYS_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    
    
   
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
                              
                
                     '//chiamo la sub con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng)
                        '
                        Call ApriFilePdf(MyPath_s, MyFile_s, 0)
                         
                    
            '//--------------------------------------------------------------------------------//--------//
                                    
              

Exit_HELP_ANALISI_ESPORTA_MSYS_Click:
    Exit Sub

Err_HELP_ANALISI_ESPORTA_MSYS_Click:
    MsgBox Err.Description
    Resume Exit_HELP_ANALISI_ESPORTA_MSYS_Click
    
End Sub



