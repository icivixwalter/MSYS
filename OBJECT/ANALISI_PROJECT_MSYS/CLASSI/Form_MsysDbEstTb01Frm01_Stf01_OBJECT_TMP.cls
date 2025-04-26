VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MsysDbEstTb01Frm01_Stf01_OBJECT_TMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'//Oggetti per il funzionamento : _
    MODULI: _
        UTIL_MDL41_01_CONTROLLO_OGGETTI_QUERY_DEL_DB _



'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
'//       LE VARIABILI DI MODULO


'//LE VARIABILI DATABASE
'//....................................................................//
    Dim daoDB As DAO.Database          '//occorre il riferimento alla libreria  Microsoft DAO 3.6 Object Libray
    Dim DaoWks As DAO.Workspace
    Dim daoRS As DAO.Recordset

    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    Dim sSql As String                          '//STRINGA SQL
    Dim MyPath_s  As String                        '//la path
    Dim MyFile_s  As String

    '//Contatori
    Dim iCount As Integer
    Dim dbl_count As Double
    
   
    'Le variabili generiche
    Dim Vv1 As Variant
    Dim Dbl1 As Double
    Dim Int1 As Integer
    Dim Long1 As Long
    Dim Str1 As String
    Dim Bool1  As Boolean
    
    '//Messaggi di errore
    Dim ProceduraMessaggioErrore_s As String    '//Errore procedura
    Dim ProceduraAttivaEseguita_s As String     '//Errore Attivita eseguita
















Private Sub Attribute_Txt_DblClick(Cancel As Integer)
'//apro gli attributi
        DoCmd.OpenQuery "MSsys_DF05_Qry01_ATTRIBUTI"

End Sub


'//@OPEN@DB_(routine per l'apertura del DATABASE)
Private Sub DATABASE_s_TXT_DblClick(Cancel As Integer)
        
        '//SOSPESO ---> MyPath_s = "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\ScannerTmp\"     ' Imposta il percorso.
        
        '//SOSPESO ---> TROVA_N01_PATH_pFunct MyPath_s
        
        
        '//CHIAMO LA PROCEDURA ED AGGIORNO LA SOTTOFORM
        '//........................................................
            '// imposta le variabile con il valore dei campi
            Dim MyFile_s As String
            Dim MyPath_s  As String
            Dim IDGestione_lng As Long
            Dim Messaggio_s As String
            
            
            MyFile_s = ""
            MyPath_s = ""
            Messaggio_s = ""
            IDGestione_lng = 0
            If Me.DATABASE_s_TXT.Value > "" Then
                MyFile_s = Me.DATABASE_s_TXT.Value
            End If
            
            If Me.DISCO_s & Me.PATH_s_TXT > "" Then
                MyPath_s = Me.DISCO_s & Me.PATH_s_TXT
                       
            End If
            
            If Me.ID_TAB_lng > "" Then
            IDGestione_lng = Me.ID_TAB_lng
            
                          
        
            '//CHIAMA LA PROCEDURA PRI FILE - @APRI.FILE
            '//-----------------------------------------------------------------------------------------------//
            '//apri file par1 = file; par2 = path
            '//ATTENZIONE occorre inserire uno spazio finale altrimenti non riconosce il percorso. Se invece è _
               un programma oppure un eseguibile non occorre lo spazio. ES _
               "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\CIVILIA_GESTIONE\CIVILIA_TEXT_ANALISI\CIVILIA_ANALISI_SUBLIME_TEXT\PROGETTO_SUBLIME_TEXT\ "
               
                '//---------------------------------------------------------------------------------------------------mantenere lo spazio ...
                'ApriFile_Sub "CIVILIA_SUBLIME_TEXT_ANALISI.sublime-project", "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\Project_PROGETTI_MDB.sublime-project "
                  
            
            '//-----------------------------------------------------------------------------------------------//
            
                        
                    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                    '//--------------------------------------------------------------------------------//--------//
                    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
                    '//CODICE               -> Function ApriFile_Pfunct.01.01
                    '//PARAMETRI            -> par_Path_s         = PATH _
                                            -> par_NameFile_s               = NOME FILE _
                                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
                            '//imposto i parametri - LA PATH
                            'MyPath_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\"
                            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
                            'MyFile_s = "Project_PROGETTI_MDB.sublime-project "
                                
                                
                             '//chiamo la sub con i parametri =   'CALL (ApriFile_Pfunct Path_programma_s, NameFile_s, IDGestione_lng)
                                '
                                Call ApriFile_Pfunct(MyPath_s, MyFile_s, 0)
                                 
                            
                    '//--------------------------------------------------------------------------------//--------//
                                            
            
           
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = ""
            ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s

                        
            End If
            
            '//Imposta le variabili a null
            Str1 = vbNull
            MyFile_s = vbNull
            MyPath_s = vbNull
            IDGestione_lng = vbNull
            
            
        '//CHIAMO LA PROCEDURA ED AGGIORNO LA SOTTOFORM *** FINE ***
        '//........................................................
   



End Sub

'//@APRO@FILE_(apro il db che contiene l'oggetto)
Private Sub Name1_TXT_DblClick(Cancel As Integer)

'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFile_Pfunct.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
        
        '//imposto i parametri - LA PATH
        '


        
'//--------------------------------------------------------------------------------//--------//
   




    On Error GoTo Err_Name1_TXT_DblClick


        '//@CONTROLLO@PATH@FILE_(prima di chiamare la routine faccio il controllo della path)
        '//-------------------------------------------------------------------------//
        MyPath_s = Me.PATH_s_TXT
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        '
        MyFile_s = Me.DATABASE_s_TXT
    
        If MyPath_s = "" Or MyFile_s = "" Then
            MsgBox "MANCA LA PATH O IL FILE USCITA DALLA ROUTINE!"
            GoTo Exit_Name1_TXT_DblClick
        End If
                    
        '//-------------------------------------------------------------------------//
        
        '//chiamo la sub con i parametri =   'CALL (ApriFile_Pfunct Path_programma_s, NameFile_s, IDGestione_lng)
        '
        Call ApriFile_Pfunct(MyPath_s, MyFile_s, 0)
         
        
            
'USCITA ED ERRORI
'..............................................................
Exit_Name1_TXT_DblClick:
    Exit Sub

Err_Name1_TXT_DblClick:
    MsgBox Err.Description
    Resume Exit_Name1_TXT_DblClick

                                                      
End Sub


'//-----------------------------------------------------------------------//






'....................................................................

'//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//




Private Sub Scel_b_AfterUpdate()
    '//NON E'AMMESSA LA SELEZIONE SE il campo TIPO OGGETTO = NULL _
       ATTENZIONE occorre il tipo oggetto valorizzato altrimenti non puoi AGGIORNARE IL CAMPO SCELTO _
       @blocco.Scel_b.record.nullo; SE ESISTE ALMENO UN RECORD viene attivata la selezione _
       vedi codice @SELEZIONA.TRUE.TUTTO.MSys_QUERY.Scel_b
    If IsNull(Me.TIPOGGETTO_s) = True Then Me.Undo
    
End Sub




Private Sub Scel_b_DblClick(Cancel As Integer)
    
        
    '//SELEZIONA COLLETTIVA MSys_TABELLE
    '//--------------------------------------------------------------------------------//-------//
    '//note : seleziona tutto a True se gia impostata altrimente imposta tutto a False _
              a condizione che la tabella è valorizzata per cui non è ammessa la selezione _
              se il campo TIPO_OGGETTO  è a null; occorrono almeno 1 record presente _
              vedi QUESTA CHIAVE ----> @blocco.Scel_b.record.nullo. _
              @SELEZIONA.TRUE.TUTTO.MSys_TABELLE.Scel_b
        
        
        If Me.Scel_b = False Then
        DoCmd.RunSQL "UPDATE " & _
                     "MSys_TABELLE " & _
                     "SET " & _
                     "MSys_TABELLE.Scel_b = True " & _
                     "WITH OWNERACCESS OPTION;"
                     
                                 
                     Me.Refresh
                    '//refresh + recalc = perche altrimenti blocca il record corrente
                    Me.Recalc
                    
        End If
        
        If Me.Scel_b = True Then
        DoCmd.RunSQL "UPDATE " & _
                     "MSys_TABELLE " & _
                     "SET " & _
                     "MSys_TABELLE.Scel_b = false " & _
                     "WITH OWNERACCESS OPTION;"
                     '//refresh + recalc = perche altrimenti blocca il record corrente
                     Me.Refresh
                     Me.Recalc

        End If
    '//--------------------------------------------------------------------------------//-------//
    
    

End Sub
