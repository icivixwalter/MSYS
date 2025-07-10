Private_Sub_Help_Txt_AfterUpdate.md


        
'//ATTIVO IL COMANDO DI GESTIONE DEL CALENDARIO
   '//========================================================================================================================================//
   '//Tipo           : Routine pubblica.
   '//Attività       : Aggiornamento casella combinata
   '//Note           : attivo il comando di aggiornamento
   '//Parametro      :
   '//Codice         : Private Sub Help_Txt_AfterUpdate().01
   '//

   Private Sub Help_Txt_AfterUpdate()

   '//MessaggiDiErrore
   Dim ProceduraMessaggioErrore_s As String
   Dim ProceduraAttivaEseguita_s As String


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
    Dim par_Help_s As String
    
                                     
                              

             
     '//....
   On Error GoTo Exit_Err_Help_Txt_AfterUpdate


         
         '//Imposto i parametri
         ProceduraAttivaEseguita_s = "Private Sub Help_Txt_AfterUpdate()"
         ProceduraMessaggioErrore_s = "Errore nella procedura"
         
     '//APRO LA TABELLA CON LA QUERY
     '//.....................................................................................................
     '//Note           : @APRI@TABELLA@HELP_(@evento @doppio@click per aprire la tabella help)

         DoCmd.OpenQuery ("PRES3000_QryDF20_01_Help_Calendario")


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

                                     par_Help_s = Me.Help_Txt
                                     par_GIORNO_CALEND_s = GiornoCalend_s

                                     
                                                                             
                                 '//----------------------------------------------------------------------------------------//
                 '//ITERO NELLA TABELLA
                 '//.....................................................................................................


                     '//RECUPERO PARAMETRO DA TABELLA OGGETTI
                     Set DaoRs = CurrentDb.OpenRecordset("PRES3000_Tb01_Calendario")
                     '//controllo se il rs è popolato'
                     If DaoRs.EOF = False And DaoRs.BOF = False Then

                             DaoRs.MoveFirst

                                 While Not DaoRs.EOF
                                     '//controllo se i campi a video sono = ai campi del rs'
                                     If DaoRs.Fields("HELP_COD_s") = Me.Help_Txt Then

                                         '//Salvo il comando da eseguire il file da aprire con la path
                                         par_Help_s = Me.Help_Txt

                                         PathFile_s = DaoRs.Fields("PathFile_s")
                                         FileCalend_s = DaoRs.Fields("FileCalend_s")

                                         
                                         '//vado all'ultimo record per chiudere il rs
                                         DaoRs.MoveLast

                                     End If

                                 DaoRs.MoveNext

                                 Wend

                             DaoRs.Close

                             Set DaoRs = Nothing

                     End If

                 '//*** fine ***
                 '//ITERO NELLA TABELLA
                 '//.....................................................................................................


                  Select Case Me.Help_Txt
                     Case "Help_Calend"

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


      Exit_Err_Help_Txt_AfterUpdate:
                           Exit Sub

      Err_Err_Help_Txt_AfterUpdate:
           MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
           Debug.Print ProceduraMessaggioErrore_s
           Debug.Print ProceduraAttivaEseguita_s
           Stop
           Resume Exit_Err_Help_Txt_AfterUpdate

   End Sub

   '//*** FINE ***
   '//ATTIVO IL COMANDO DI GESTIONE DEL CALENDARIO
   '//========================================================================================================================================//





