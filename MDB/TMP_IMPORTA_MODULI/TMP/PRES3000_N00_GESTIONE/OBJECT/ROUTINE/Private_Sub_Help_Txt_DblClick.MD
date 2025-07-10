Private_Sub_Help_Txt_DblClick.MD
    
    Note
        apro la tabella help

    SCHEMA

        '//APRI TABELLA HELP
            '//========================================================================================================================================//
            '//Tipo           : Routine pubblica.
            '//Attività       : Apri tabella 
            '//Note           : Apro la tabella help della gestione del calendario per visualizare gli elementi e per la
            '//Parametro      : cancel  = reset a zero
            '//Codice         : Private Sub Help_Txt_DblClick(Cancel As Integer).01
            '//

            Private Sub Help_Txt_DblClick(Cancel As Integer)

            '//MessaggiDiErrore
            Dim ProceduraMessaggioErrore_s As String
            Dim ProceduraAttivaEseguita_s As String
             
             
            '//Campo
            Dim CampoCercato_s As String

            '//Campi parametri
            Dim par_AnnoImp_i As Integer
            Dim par_CodiceTributo_s As String

                        
                '//....
            On Error GoTo Exit_Err_Help_Txt_DblClick


                    
                    '//Imposto i parametri
                    ProceduraAttivaEseguita_s = "Private Sub Help_Txt_DblClick(Cancel As Integer)"
                    ProceduraMessaggioErrore_s = "Errore nella procedura"
                    
                '//APRO LA TABELLA CON LA QUERY
                '//.....................................................................................................
                '//Note           : @APRI@TABELLA@HELP_(@evento @doppio@click per aprire la tabella help)

                    Docmd.OpenQuery("PRES3000_QryDF20_01_Help_Calendario")

              

                '//*** fine ***
                '//APRO LA TABELLA CON LA QUERY
                '//.....................................................................................................

            '//USCITA  E GESTIONE ERRORI
            '//..............................................................................................................


            Exit_Err_Help_Txt_DblClick:
                Exit Sub

            Err_Err_Help_Txt_DblClick:
                MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
                Debug.Print ProceduraMessaggioErrore_s
                Debug.Print ProceduraAttivaEseguita_s
                Stop
                Resume Exit_Err_Help_Txt_DblClick

            End Sub

            '//*** FINE ***
            '//APRI TABELLA HELP
            '//========================================================================================================================================//
