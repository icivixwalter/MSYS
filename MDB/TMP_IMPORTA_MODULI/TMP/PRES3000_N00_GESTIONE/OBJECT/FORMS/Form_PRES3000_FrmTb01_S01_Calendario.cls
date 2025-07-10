VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PRES3000_FrmTb01_S01_Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


'//@modulo_(DELLA FORM @CALENDARIO denominato in: PRES3000_Frm01_S01_01_Calendario)
'//dove si trova il modulo, bisogna aprire il progetto mdb denominato:
'
'   START "Apri proggetto Calendario" call "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N01_CALENDARIO\PRES3000_N01_CALENDARIO.mdb"

' L'ANALISI del modulo si trova qui:
'    START "Apri proggetto Calendario" call "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N00_GESTIONE\OBJECT\MODULI\PRES3000_Frm01_S01_01_Calendario.MD"
'//faq: _
    @trova@classe@s01@calendario _
    @modifica@giustificazioni@timbrature


'DAO
Dim DaoDB As DAO.Database
Dim DaoWks As DAO.Workspace
Dim DaoRs As DAO.Recordset
Dim DaoRs_Lett As DAO.Recordset

'ADO
Dim ADODB As Database
Dim AdodaoRs As Recordset

'Contatori
Dim iCount As Integer
Dim iTOTcount As Integer
Dim dbl_count As Double

'Le variabili generiche
Dim sSql As String                                          ' Stringa di estrazione
Dim sSql_Lett As String                                     ' Stringa di estrazione

'Variabili generali
Dim Str1 As String
Dim Int1 As Integer
Dim Int2 As Integer
Dim Int3 As Integer
Dim Lng1 As Long
Dim Dbl1 As Double
Dim Bln1 As Boolean
Dim Vv1 As Variant
Dim obj1 As Object

'Gestione parametri comandi
Dim sxParamCmd_3 As String


'Larghezza e numero di colonna
Dim sLarg_Col As String
Dim iNum_Col As Integer


'//ERRORI PROCEDURA_FUNCTION O ROUTINE
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String



'//CAMPO FILE CALENDARIO EVENTO
'//========================================================================================================================================//
'//Tipo           : Routine Privata.
'//Attività       : Aggiornamento casella combinata
'//Note           : @file@calendario@aggiornamento AGGIORNO la proprieta RecordoSource
'//Parametro      :
'//Codice         : FILE_CALEND_Txt_GotFocus().01
'//



'//@CARICO@SQL_(NEL CAMBO FILE
Private Sub FILE_CALEND_Txt_GotFocus()

'//i parametri giorno
Dim Anno_i As Integer
Dim Mese_s As String
Dim Giorno_s As String
Dim Stringa_s As String
Dim TipoDescr_s As String

    
    

     '//....
   On Error GoTo Exit_Err_FILE_CALEND_Txt_GotFocus

         
         '//Imposto i parametri Messaggio di errore
         ProceduraAttivaEseguita_s = "FILE_CALEND_Txt_GotFocus()"
         ProceduraMessaggioErrore_s = "Errore nell'impostazione della proprieta RecordSource"
         
 
    
    
    '//@RECORDSOURCE_(DEL CAMPO TABELLA @FILE@CALEND)
    '//----------------------------------------------------------------------------------------
    '//Note     : @imposto@RecordSource_(LA proprieta recordSource del campo FILE CALENDARIO)
    '//CODICI   :  @file@calendario@impostazioni, @proprieta@recordSouce@campo@FILE_CALEND_Txt
    '//          @REIMPOSTA@PARTE@FINALE_(del giustificativo, reimpostando i codici)

    
        '//RESET DELLE VARIABILI
           Anno_i = 0
           Giorno_s = ""
           Mese_s = ""
           Stringa = ""
           TipoDescr_s = ""
        
           '//@REIMPOSTO@CODICI_(vengono reimpostati i codici secondo quanto previsto nella @SELECT)
           '//Imposto la variabile descrizione ed eventualmente modifico il tipo di descrizione
           TipoDescr_s = "_WALTER_{GG_1}.BAT"
           Select Case Me.COD_BREV_GIORN_Txt
           
                '//-----------------TIMBRATURE ---------------------------
                '//@reimposto il codice Badge Dimenticato
                Case "BadgeD"
                    TipoDescr_s = "_WALTER_{E_08_00-U_14_20}.BAT"
                    
                '//@reimposto il codice @TIMBRATURA ERRATA solo la parte finale
                Case "TimbErr"
                    TipoDescr_s = "_WALTER_{E_08_00}.BAT"
           
                
                '//@reimposto il codice @TIMBRATURA FUORI SEDE solo la parte finale
                Case "TimbFSed"
                    TipoDescr_s = "_WALTER_{E_08_00_U_14_00}_{Sede_MTC}.BAT"
                
                
                
                '//@reimposto il codice @USCITA ANTICIPATA A RECUPERO solo la parte finale
                Case "UscitaAnticipat"
                    TipoDescr_s = "_WALTER_{U_14_00}.BAT"
                
               
                '//@reimposto il codice @TIMBRATURA @DA @CORREGGERE solo la parte finale
                Case "ErroreTimbratur"
                    TipoDescr_s = "_WALTER_{U_14_00}.BAT"
               
               
               
               
                '//-----------------CARTELLINI ---------------------------
                
                   
                '//@reimposto il codice @CARTELLINO solo la parte finale
                Case "Cartellino"
                    TipoDescr_s = "_MESE_{_" & Me.MESE_CALEND & "_}.BAT"
            
                
                
                '//-----------------PERMESSI ---------------------------
                '//@reimposto il codice @IncrementoOrario
                 Case "IncremOrario"
                    TipoDescr_s = "_WALTER_{E_08_00-U_14_20}_{ORE_00_00}.BAT"
                '//@reimposto il codice @P.breve solo la parte finale
                 Case "P.Breve"
                    TipoDescr_s = "_{DALLE_ORE_00_00-ALLE_00_00}.BAT"
                '//@recupero@permesso@preve al posto dell'ex @incremento@orario @p.breve
                 Case "RecuperoP.Breve"
                    TipoDescr_s = "_{DALLE_ORE_00_00-ALLE_00_00}.BAT"
                    
                    
                '//-----------------STRAORDINARI ---------------------------
                    
                '//@reimposto il codice @STRAORDINARIO solo la parte finale
                 Case "Straord"
                    TipoDescr_s = "_{DALLE_ORE_00_00-ALLE_00_00}_{ORE_00_00}.BAT"
                      
                '//@reimposto il codice @STRAORDINARIO @ELETTORALE solo la parte finale
                 Case "StraordELETT"
                    TipoDescr_s = "_{DALLE_ORE_00_00-ALLE_00_00}_{ORE_00_00}.BAT"
                  
                    
                '//@reimposto il codice @STRAORDINARIO @FESTIVO solo la parte finale
                 Case "StraordFES"
                    TipoDescr_s = "_{DALLE_ORE_00_00-ALLE_00_00}_{ORE_00_00}.BAT"
                  
                    
           Case Else
           
           End Select
           
           Anno_i = Me.AnnoCalend_i
           Giorno_s = Day(Me.DATA_CALEND)
           If Giorno_s < 10 Then Giorno_s = "0" & Giorno_s
           
            Mese_s = Month(Me.DATA_CALEND)
           If Mese_s < 10 Then Mese_s = "0" & Mese_s
        
           Stringa = Stringa & Anno_i & "_" & Mese_s & "_" & Giorno_s & "_" & Me.DESCRIZ_GIORN_s_TXT & TipoDescr_s
           
           sSql = ""
           
           
           sSql = ""
           sSql = sSql & "SELECT "
           sSql = sSql & "'" & Stringa & "'" & " AS [STRINGA_c] ,  "
           sSql = sSql & "PRES3000_Tb01_Calendario.Cod_BrevGiorn_s        AS [COD_BREV_c], "
           sSql = sSql & "PRES3000_Tb01_Calendario.DataCalend_d,  "
           sSql = sSql & "PRES3000_Tb01_Calendario.ID,  "
           sSql = sSql & "PRES3000_Tb01_Calendario.DESCRIZ_GIORN_s        AS [DESCRIZ_GIORN_c], "
           sSql = sSql & "PRES3000_Tb01_Calendario.AnnoCalend_i           AS [ANNO_c], "
           sSql = sSql & "'" & Giorno_s & "'" & " AS [DAY_c],"
           sSql = sSql & "'" & Mese_s & "'" & " AS [MOUNT_c],"
           sSql = sSql & "'sqlQueryCalendario' AS TxtCodiceQry, "
           sSql = sSql & "'WALTER' AS WR_c "
           sSql = sSql & "FROM "
           sSql = sSql & "PRES3000_Tb01_Calendario "
           sSql = sSql & "WHERE (((PRES3000_Tb01_Calendario.ID)=" & Me.ID & "));"
           Debug.Print sSql
           
           '//inserimento della query costruita nella proprieta RecordSource del campo
           Me.FILE_CALEND_Txt.RowSource = ""
           Me.FILE_CALEND_Txt.RowSourceType = "Tabella/Query"
           Me.FILE_CALEND_Txt.RowSource = sSql
    '//----------------------------------------------------------------------------------------
         

   '//USCITA  E GESTIONE ERRORI
   '//..............................................................................................................

         
Exit_Err_FILE_CALEND_Txt_GotFocus:
Exit Sub

Err_FILE_CALEND_Txt_GotFocus:
           MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
           Debug.Print ProceduraMessaggioErrore_s
           Debug.Print ProceduraAttivaEseguita_s
           Stop
           Resume Exit_Err_FILE_CALEND_Txt_GotFocus
 End Sub
'//========================================================================================================================================//




'//ATTIVO IL COMANDO DI GESTIONE DEL CALENDARIO
   '//========================================================================================================================================//
   '//Tipo           : Routine pubblica.
   '//Attività       : Aggiornamento casella combinata
   '//Note           : @help@comando@aggiornamento AGGIORNO oppure apro i file od eseguo un comando  @evento@dopo@aggiornamento
   '//Parametro      :
   '//Codice         : Private Sub HELP_COD_s_Txt_AfterUpdate().01
   '//

Private Sub HELP_COD_s_Txt_AfterUpdate()

  
   '//Campo
   Dim CampoCercato_s As String

   '//Campi parametri
   Dim par_AnnoImp_i As Integer
   Dim par_Help_s As String
    
    Dim MyPath_s As String
    Dim MyFile_s As String
    Dim PathFile_s As String
    Dim FileCalend_s As String
    Dim IDGestione_lng As Long
    
    
                                     
                              

             
     '//....
   On Error GoTo Exit_Err_HELP_COD_s_Txt_AfterUpdate


         
         '//Imposto i parametri
         ProceduraAttivaEseguita_s = "Private Sub HELP_COD_s_Txt_AfterUpdate()"
         ProceduraMessaggioErrore_s = "Errore nella procedura"
         
     '//APRO LA TABELLA CON LA QUERY
     '//.....................................................................................................
     '//Note           : @APRI@TABELLA@HELP_(@evento @doppio@click per aprire la tabella help)




                 '//ITERO NELLA TABELLA
                 '//....................................................................................................

                 '//Note           : Tramite una Select vengono individuati i valori da restiuire.


                        '//APRO FILE RECUPERATO NELLA TABELLA
                                 '//----------------------------------------------------------------------------------------//
                                 '//NOTE                 -> Apro il file di tipo doc, zip o pdf
                                 '//CODICE               -> Function ApriFilePdf.01.01

                                     PathFile_s = ""
                                     FileCalend_s = ""
                                     IDGestione_lng = 0

                                     par_Help_s = ""
                                     par_GIORNO_CALEND_s = ""
                                     
                                  

                                     '//imposta
                                     par_Help_s = Me.HELP_COD_s
                                                                             
                                 '//----------------------------------------------------------------------------------------//

                 '//ITERO NELLA TABELLA
                 '//.....................................................................................................
                            
                            
                            sSql = ""
                            sSql = sSql & "SELECT PRES3000_DF20_Help_Calendario.HELP_COD_s, "
                            sSql = sSql & "PRES3000_DF20_Help_Calendario.DESCRIZIONE_HELP_s,"
                            sSql = sSql & "PRES3000_DF20_Help_Calendario.NOTE_HELP_s, "
                            sSql = sSql & "PRES3000_DF20_Help_Calendario.File_s,"
                            sSql = sSql & "PRES3000_DF20_Help_Calendario.PathFile_s "
                            sSql = sSql & "FROM PRES3000_DF20_Help_Calendario "
                            sSql = sSql & "WHERE (((PRES3000_DF20_Help_Calendario.HELP_COD_s)='" & par_Help_s & "'));"

                       
                            
                            '//imposto la stringa di estrazione
                            Debug.Print "controllo"
                            Debug.Print par_Help_s
                            Debug.Print Me.DataCalend_d
                            Debug.Print sSql
                            


                     '//RECUPERO PARAMETRO DA TABELLA OGGETTI
                     Set DaoRs = CurrentDb.OpenRecordset(sSql)
                     '//controllo se il rs è popolato'
                     If DaoRs.EOF = False And DaoRs.BOF = False Then

                             DaoRs.MoveFirst

                                 While Not DaoRs.EOF
                                    

                                         '//Salvo il comando da eseguire il file da aprire con la path
                                         par_Help_s = Me.HELP_COD_s_Txt

                                         PathFile_s = DaoRs.Fields("PathFile_s")
                                         FileCalend_s = DaoRs.Fields("File_s")

                                         IDGestione_lng = 0

                                         
                                         '//vado all'ultimo record per chiudere il rs
                                         'DaoRs.MoveLast

                                 DaoRs.MoveNext

                                 Wend

                             DaoRs.Close

                             Set DaoRs = Nothing

                     End If

                 '//*** fine ***
                 '//ITERO NELLA TABELLA
                 '//.....................................................................................................


                  Select Case Me.HELP_COD_s_Txt
                     Case "CALENDARIO_HELP"

                          '//apro l'Help  PASSANDO I PARAMETRI
                             Call ApriFilePdf(PathFile_s, FileCalend_s, IDGestione_lng)
                     
                     '... inserire altri case'

                     Case Else
                     
                     End Select
                         

     '//*** fine ***
     '//APRO LA TABELLA CON LA QUERY
     '//.....................................................................................................

   '//USCITA  E GESTIONE ERRORI
   '//..............................................................................................................


Exit_Err_HELP_COD_s_Txt_AfterUpdate:
Exit Sub

Err_HELP_COD_s_Txt_AfterUpdate:
           MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
           Debug.Print ProceduraMessaggioErrore_s
           Debug.Print ProceduraAttivaEseguita_s
           Stop
           Resume Exit_Err_HELP_COD_s_Txt_AfterUpdate
 End Sub




'//*** FINE ***
'//ATTIVO IL COMANDO DI GESTIONE DEL CALENDARIO
'//========================================================================================================================================//



Private Sub Help_Txt_BeforeUpdate(Cancel As Integer)
    Debug.Print
End Sub

        



'//APRI TABELLA HELP
    '//========================================================================================================================================//
    '//Tipo           : Routine pubblica.
    '//Attività       : Apri tabella
    '//Note           : Apro la tabella help della gestione del calendario per visualizare gli elementi e per la
    '//Parametro      : cancel  = reset a zero
    '//Codice         : Private Sub HELP_COD_s_Txt_DblClick(Cancel As Integer).01
    '//

Private Sub HELP_COD_s_Txt_DblClick(Cancel As Integer)

            '//MessaggiDiErrore
            Dim ProceduraMessaggioErrore_s As String
            Dim ProceduraAttivaEseguita_s As String
             
             
            '//Campo
            Dim CampoCercato_s As String

            '//Campi parametri
            Dim par_AnnoImp_i As Integer
            Dim par_CodiceTributo_s As String

                        
                '//....
            On Error GoTo Exit_Err_HELP_COD_s_Txt_DblClick


                    
                    '//Imposto i parametri
                    ProceduraAttivaEseguita_s = "Private Sub HELP_COD_s_Txt_DblClick(Cancel As Integer)"
                    ProceduraMessaggioErrore_s = "Errore nella procedura"
                    
                '//APRO LA TABELLA CON LA QUERY
                '//.....................................................................................................
                '//Note           : @APRI@TABELLA@HELP_(@evento @doppio@click per aprire la tabella help)

                    DoCmd.OpenTable ("PRES3000_DF20_Help_Calendario")

              

                '//*** fine ***
                '//APRO LA TABELLA CON LA QUERY
                '//.....................................................................................................

            '//USCITA  E GESTIONE ERRORI
            '//..............................................................................................................


Exit_Err_HELP_COD_s_Txt_DblClick:
                Exit Sub

Err_HELP_COD_s_Txt_DblClick:
                MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
                Debug.Print ProceduraMessaggioErrore_s
                Debug.Print ProceduraAttivaEseguita_s
                Stop
                Resume Exit_Err_HELP_COD_s_Txt_DblClick

End Sub



'//*** FINE ***
'//APRI TABELLA HELP
'//========================================================================================================================================//









'//*** fine ***
'//Help_Txt_Change()
'//================================================================================================//





Private Sub NOTE_GIORN_txt_AfterUpdate()

'//AGGIORNA DATA E ORA
'//...................................................................//
'//
    Call AGGIORNA_DATA_ORA_Sub

'//...................................................................//

End Sub


Private Sub Indirizzo_DblClick(Cancel As Integer)

    '//Aggiorna l'indirizzo esteso
    DoCmd.OpenQuery "A0203_Qry02-42_(IndirizzoEsteso)_(Up)"
    


End Sub

'//EVENTI @CODICE@BREVE@GIORNALIERO
'//========================================================================//
Private Sub COD_BREV_GIORN_Txt_AfterUpdate()
    '//@Aggiornamento_(con il codice breve aggiorno il campo @descrizione @codice@breve)
    Debug.Print "codice breve: " & Me.COD_BREV_GIORN_Txt.Column(0) & " descrizione codice breve: "; Me.COD_BREV_GIORN_Txt.Column(1)
    Me.DESCRIZ_GIORN_s_TXT = Me.COD_BREV_GIORN_Txt.Column(1)
End Sub

Private Sub COD_BREV_GIORN_Txt_DblClick(Cancel As Integer)
    '//APRO_TABELLA_DEF_CODICI = codici giornalieri
    DoCmd.OpenTable "PRES3000_DF10_GIORNALIERE"
End Sub


'//========================================================================//

Private Sub DataCalend_Txt_AfterUpdate()

'//AGGIORNA DATA E ORA
'//...................................................................//
'// CALL
        AGGIORNA_DATA_ORA_Sub

'//...................................................................//

End Sub











'//CERCO IL FILE
Private Sub FILE_CALEND_Txt_DblClick(Cancel As Integer)



    '//Ricerca del file solo se il campo è valorizzato
    '//If Me.FILE_CALEND_Txt > "" Then
    '//..........................................................................
        If Me.FILE_CALEND_Txt > "" Then
            
                
                '//If Me.PATH_TXT > "" Then
                '//Note : Controllo se il campo path è valorizzato
                '//..........................................................................
                      
                    If Me.PATH_TXT > "" Then
                                            
                        '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                        '//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                        '//Note: Apro il file PDF speficico indicando come paramentro la path completa
                        '//     incluso il file PDF.
                                           
                                '//Definisco ed imposto la variabile di ricerca + reset variabili
                                '// imposta le variabile con il valore dei campi
                                  Dim MyFile_s As String
                                  Dim MyPath_s  As String
                                  Dim IDGestione_lng As Long
                                    
                                    MyFile_s = ""
                                    MyPath_s = ""
                                    IDGestione_lng = 0
                                    
                                    MyFile_s = Me.FILE_CALEND_Txt.Value
                                    MyPath_s = Me.PATH_TXT.Value
                                    IDGestione_lng = Me.ID
                                    IDGestione_lng = 0
                                    
                                    
                                    
                                    Debug.Print "controllo variabili -> MyPath_s " & Path_programma_s
                                    Debug.Print "controllo variabili -> MyFile_s  " & NameFile_s
                                    
                                    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                                    '//----------------------------------------------------------------------------------------//
                                    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
                                    '//CODICE               -> Function ApriFilePdf.01.01
                                    
                                             '//chiamo la sub con i parametri
                                                Call ApriFilePdf(MyPath_s, MyFile_s, IDGestione_lng)
                                                                                
                                    '//----------------------------------------------------------------------------------------//

                                    
                                    '//reset variabili
                                    Path_programma_s = vbNull
                                    NameFile_s = vbNull
                                    Path_s = vbNull

                                
                          '//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  
                    
                    
                        
                        Else
                            MsgBox "PATH del file non valorizzata - Ricerca Sospesa", vbCritical, "MSG_PATH_FILE"
                            
                    End If
                '//If Me.PATH_TXT > "" Then *** fine ***
                '//..........................................................................
                
           End If
                
    '//If Me.FILE_CALEND_Txt > "" Then *** fine ***
    '//..........................................................................
    
    
    
End Sub




'//APRI_IL_FILE_PDF
'//==================================================================================================================//
'//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
'//PARAMETRI        : passo 2 stringhe per parametro, la path e il nome del file.pdf

Public Sub ApriFilePdf(par_Path_programma_s As String, par_NameFile_s As String, par_IDGestione_lng As Long)

Dim Path_programma_s As String
Dim NameFile_s As String
Dim Stringa1 As String

On Error GoTo ApriFilePdf_Err
        
        
                      '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                        '//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                        '//Note: Apro il file PDF speficico indicando come paramentro la path completa
                        '//     incluso il file PDF.
                                           
                                '//Definisco ed imposto la variabile di ricerca + reset variabili
                                Dim MyPath_s As String
                                Dim MyFileEstensione_s As String
                                
                                    '// Path_programma_s = Me.PATH_TXT.Value
                                    '// NameFile_s = Me.FILE_CALEND_Txt.Value
                                    '// Debug.Print "controllo variabili -> Path_programma_s " & Path_programma_s _
                                    Debug.Print "controllo variabili -> NameFile_s " & NameFile_s
                                    
                                    '//chiamo la procedura
                                    '//Call ApriFilePdf(Path_programma_s, NameFile_s)
                                    
                                    '//reset variabili
                                    Path_programma_s = vbNull
                                    NameFile_s = vbNull
                                    Path_s = vbNull
                                   

                                

                                '//TROVA FILE  (INIZIO)_E_RECUPERO_LA_ESTENSIONE
                                '//------------------------------------------------------------------------------------------------
                                    ' Visualizza i nomi in c:\ che rappresentano directory.
                                     MyPath_s = par_Path_programma_s
                                
                                    Debug.Print "                       CONTROLLO FILE TROVATO                      "
                                    Debug.Print "..................................................................."
                                    Debug.Print "Path -> " & MyPath_s
                                            
                                            MyFile_s = Dir(par_Path_programma_s & par_NameFile_s & ".*", vbNormal)
                                            If MyFile_s > "" Then
                                                Debug.Print MyFile_s              ' Visualizza la voce solo
                                                
                                                
                                                '//SOSPESO _
                                                 MsgBox "OK_CONTROLLO_FILE=TROVATO--->" & Chr$(13) & MyPath_s & Chr$(13) & MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
                                                
                                                
                                                
                                                '//TROVA L'ESTENSIONE _
                                                Note : Esamina gli ulti 4 caratteri del file e li salvo su una variabile _
                                                che sarà usato per completare l'intera path+nomefile da applicare all'oggetto _
                                                di seguito richiamato con la shell. Il salvataggio avviene solo se tra il _
                                                file passato come parametro e quello ricercato mediante dir vi è una differenza. _
                                                Se vi è una differenza allora significa che probabilmente è senza estensione.
                                                If MyFile_s <> par_NameFile_s Then
                                                    MyFileEstensione_s = Right(MyFile_s, 4)
                                                End If
                                                
                                            Else
                                                
                                                '//Se non trova il file emette il messaggio ed esco dalla routine e scrivo sul campo _
                                                NoteAtto_s = "FILE_NON_TROVATO", mediante la query di aggiornamento basata sull'id del record.
                                                
                                                '//Aggiorno la nota della tabella con la dicitura file non trovato se esiste Id del record
                                                If par_IDGestione_lng > 0 Then
                                                    '//Controllo ed esecuzione cmd_sSql
                                                    sSql = ""
                                                    sSql = sSql & "UPDATE LLPP_ATTI_Tb01_Gestione SET LLPP_ATTI_Tb01_Gestione.RICERCA_FileAtto_s = 'FILE TROVATO'"
                                                    sSql = sSql & "WHERE (((LLPP_ATTI_Tb01_Gestione.IDGestione)=" & par_IDGestione_lng & "));"
                                                    Debug.Print sSql
                                                    
                                                    '//SALVATAGGIO SOSPESO
                                                    CurrentDb.Execute sSql
                                                    
                                                End If
                                                
                                                Debug.Print "FILE NON TROVATO"
                                                MsgBox "ATTENZIONE!!! FILE_NON_TROVATO--->" & Chr$(13) & MyPath_s & "/" & par_MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
                                                'Me.Recalc
                                                
                                                '//libero la memoria
                                                MyFileEstensione_s = vbNull
                                                MyPath_s = vbNull
                                                par_IDGestione_lng = vbNull
                                                GoTo ApriFilePdf_Exit
                                            End If
                                    
                                    Debug.Print "..................................................................."
                                    
                                   ' MsgBox "FILE_TROVATO--->" & MyPath_s & "/" & par_MyFile_s, vbInformation, "MSG_BOX_DI_AVVISO"
                                     
                                
                                '//TROVA FILE__E_RECUPERO_LA_ESTENSIONE  (***FINE***)
                                '//------------------------------------------------------------------------------------------------

                                
                                
                                
                          '//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
         
  '//APRO FILE PDF SPECIFICO
  '//_______________________________________________________________________________________
  '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
  '//           incluso il file PDF.
    
    '//La path del file = Unisco la path + il file+estensione sopra individuata
    par_Path_programma_s = par_Path_programma_s & par_NameFile_s & MyFileEstensione_s
           
    '//controllo
    Debug.Print par_Path_programma_s
           
    '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
    '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
    '// prima la sessione dos e poi il pdf..
    '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
    Set WshPDF = CreateObject("wscript.shell")
    
    '// Unisco il comando pdf start + path definitiva
    Stringa1 = "%comspec% /c start " & par_Path_programma_s                               '//Comando di apertura
    '//Attivo il comando
    WshPDF.Run Stringa1
    
    '//libero la memoria dalle variabili e dagli oggetti creati
    WshPDF = Null
    Stringa1 = vbNull
    MyFileEstensione_s = vbNull
    MyPath_s = vbNull
'_______________________________________________________________________________________


      
      
ApriFilePdf_Exit:
    Exit Sub

ApriFilePdf_Err:
    MsgBox Error$
    Resume ApriFilePdf_Exit

End Sub

'//APRI_IL_FILE_PDF         *** FINE ***
'//==================================================================================================================//




'//************************************************************************************************//
'//                 LE ROUTINE PUBBLICHE
'//************************************************************************************************//

Private Sub AGGIORNA_DATA_ORA_Sub()


'//AGGIORNA DATA E ORA
'//...................................................................//
'// CALL AGGIORNA_DATA_ORA_Sub

'//...................................................................//
    
    '//AGGIORNA LA DATA DI AGGIORNAMENTO E L'ORA
    '//...................................................................//
        Me.DATAAGG_TXT = Date
        Me.TIMEOPER_TXT = Now()
        
        
        '//@aggiorno@mese
        Me.MeseCalend_i = Month(Me.DataCalend_Txt)
        '//@aggiorno@il@giono@della@settimana estranedo il giorno dal lunedi alla domenica con Format (.....)
        Str1 = Format(Me.DataCalend_Txt, "dddd")            '//@estraggo@giorno
        Debug.Print "GIORNO .... " & Str1
        '//aggiorno il giorno
        Me.GiornoCalend_s_Txt = Str1
        
        '//aggiorno l'anno @AGGIORNO@ANNO '//@aggiorno@anno
        Me.AnnoCalend_i_TXT = Year(Me.DataCalend_Txt)
          
        
        
    '//...................................................................//
 


End Sub



'// AGGIORNAMENTO CAMPO CALENDARIO
'//-----------------------------------------------------------------------------------------------//
'//ATTIVITA         : Aggiorna il campo calendario con una stringa precostruita in formato
'//                   anno mese giorno in base a questo modello, AAAMMGG_WALTER_ATTIVITA_{GG.1}
'//CODICE           : FILE_CALEND_Txt_AfterUpdate_Sub.00



Private Sub FILE_CALEND_Txt_AfterUpdate_Sub()
 '//MessaggiDiErrore
   Dim ProceduraMessaggioErrore_s As String
   Dim ProceduraAttivaEseguita_s As String


    On Error GoTo Exit_Err_FILE_CALEND_Txt_AfterUpdate_Sub


      
         '//Imposto i parametri
         ProceduraAttivaEseguita_s = "FILE_CALEND_Txt_AfterUpdate_Sub()"
         ProceduraMessaggioErrore_s = "Errore nella procedura"
    
    
    
    '//ATTIVITA
    '//---------------------------------------------------------------------------------//
    '//Note             : esecuzione ...
    '//CODICE           : FILE_CALEND_Txt_AfterUpdate_Sub.01
        
        Dim AAAA_s As String
        Dim GG_s As String
        Dim MM_s As String
        Dim ATTIVITA_s As String
            AAAA_s = Me.Anno_Key
            GG_s = Me.MESE_CALEND
            GG_s = Me.GIORNO_CALEND
            
            Debug.Print "controllo " & "AAAA_s , GG_s ,  gg_s"
        
            '//IMPOSTO le stringhe AAAA GG MM
    
    
    '//---------------------------------------------------------------------------------//
    
    
    

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_Err_FILE_CALEND_Txt_AfterUpdate_Sub:
                Exit Sub

Err_FILE_CALEND_Txt_AfterUpdate_Sub:
                       MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
                       Debug.Print ProceduraMessaggioErrore_s
                       Debug.Print ProceduraAttivaEseguita_s
                       Stop
                       Resume Exit_Err_FILE_CALEND_Txt_AfterUpdate_Sub
End Sub
'//-----------------------------------------------------------------------------------------------//






'//@blocco_(dopo l'aggiornamento del record viene controllollato ed attivato il blocco record)
'//@AGGIORNAMENTO@RECORD_(@Routine privata per @evento Form_BeforeUpdate @prima del salvataggio della Form)
'//-----------------------------------------------------------------------------------------------//
Private Sub Form_BeforeUpdate(Cancel As Integer)

Dim BLOCC_b As Boolean

 '//AGGIORNAMENTO RECORD - BLOCCO RECORD + UNDO - REIMPOSTO I VALORI PRECEDENTI
 '//........................................................
    '//Nota = nel caso dell'aggiornamento del record, deve _
      essere evidenziato un avviso con il quale si richiede lo sblocco del record _
      nel caso in cui il campo BLOCC_b = TRUE. Manualmente impostare a false e poi _
      effettuare l'aggiornamento - VIENE UTILIZZATO IL METODO UNDO  che non consente _
      l'aggiornamento del record e ripristina i valori precedenti, solo nel caso di _
      BLOCCO_RECORD = TRUE.

        '::::
            BLOCC_b = Me.BLOCC_b_Txt
            
            If BLOCC_b = True Then
                   
                  
                    
                        MsgBox "PER L'AGGIORNAMENTO TOGLIERE IL BLOCCO RECORD! E RIPETERE L'OPERAZIONE", vbExclamation, "BLOCCO DEL RECORD"
                        '//REIMPOSTA I VALORI ALLO STATO PRECEDENTE + imposta il blocco record
                        Me.Undo
                        '//Imposto il blocco record dopo l'evento unto per consentire l'aggiornamento della _
                           casella combinata altrimenti se viene inserito prima dell'evento Undo, l'aggiornamento _
                           cancella tutte le modifiche del record.
                      Me.BLOCC_b_Txt = True
                   
                
               
                
            End If
    '//........................................................



End Sub
'//-----------------------------------------------------------------------------------------------//



'//Variabili di database e GENERALI     *** FINE ***
'//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Private Sub GiornoCalend_s_Txt_AfterUpdate()
    

'//AGGIORNA DATA E ORA
'//...................................................................//
'//
    Call AGGIORNA_DATA_ORA_Sub

'//...................................................................//



End Sub
