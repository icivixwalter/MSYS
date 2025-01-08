MENU COME IMPOSTARE UN ELENCO DI MENU
    
    Nota
        Come impostare un menu ad elenco a discesa in una casella combinata occorre per 
        prima cosa gestire il reset della casella combinata nell'evento Open della form, poi
        gestire l'evento focus della casella combinata con il caricamento del menu e poi
        nell'evento SU MODIFICA della casella combinata dove viene chiama la funzione di
        scelta del menu che si trova nel modulo di gestione dei menu della casella combinata
        denominato:
                CASELLA_COMBINATA_Mdl_01_GestioneMENU


            01 SU CARICAMENTO DELLA FORM EVENTO LOAD RESET COMBINATA

                '//@EVENTO OPEN FORM IMPOSTO LA COMBINATA
                '//##################################################################################################################//
                Private Sub Form_Load()

                On Error GoTo Err_Form_Load

                        '//SU CARICAMENTO DELLA FORM RESETTO LA CASELLA COMBINATA
                        Me.COMANDO_s_TXT.Value = ""
                        Me.COMANDO_s_TXT.Requery




                '//CHIUSURA E CONTROLLO ERRORI
                '//-----------------------------------------------------------------------------------------//
                Exit_Form_Load:
                    Exit Sub

                Err_Form_Load:
                    MsgBox Err.Description
                    Stop
                    Resume Exit_Form_Load

                End Sub


                '//@EVENTO OPEN FORM IMPOSTO LA COMBINATA
                '//##################################################################################################################//

            02 IMPOSTARE L'EVENTO FOCUS DELLA CASELLA COMBINATA

                    '//@COMBINATA_Help_TXT_GotFocus
                    '//##################################################################################################################//
                    Private Sub Help_TXT_GotFocus()

                    On Error GoTo Err_Help_TXT_GotFocus




                        
                        '//IMPOSTO@LA@COMBINATA AL FOCUS DELLA CASELLA @IMPOSTA@MENU @MENU@PAG.12  @MENU.DELLA.PAGINA.12
                        '//------------------------------------------------------------------------------------------------------//
                        '//NOTE Se la proprietà RowSourceType è impostata a Elenco valori, _
                          le impostazioni valide per le proprietà RowSource e ColumnCount e gli elenchi che ne risultano saranno: _
                           Me.Help_TXT.ColumnCount = 1 AD UNA COLONNA; _
                           Me.Help_TXT.RowSource = "1_Lun;2_Mar;3_Mer;"  CON 3 RIGHE. Viene impostata sull'attivazione o focus
                            
                            '//imposto il valore base
                            Me.Help_TXT.Value = "MENU_INIZIALE"
                            
                          
                            '//imposto ad elenco valori
                            Me.Help_TXT.RowSourceType = "Elenco valori"
                            
                            '// 1 COLONNA RIGHE 3 COLUME,ROW
                            Me.Help_TXT.ColumnCount = 1
                            '//imposto i valori della casella combinata su 3 righe 1 COLONNA
                            Me.Help_TXT.RowSource = "1_ANALISI_PROGETTO_ANAGRAFICA;" & _
                                                     "2_Mar;" & _
                                                     "10_OGGETTI_MSys_ApriProgettiTutti;" & _
                                                     "20_HELP_PROGETTI_ApriPdf;"
                            
                            
                            '//riaggiorna
                            Me.Help_TXT.Requery
                            
                            
                            
                            
                          
                        '//------------------------------------------------------------------------------------------------------//


                    '//CHIUSURA E CONTROLLO ERRORI
                    '//-----------------------------------------------------------------------------------------//
                    Exit_Help_TXT_GotFocus:
                        Exit Sub

                    Err_Help_TXT_GotFocus:
                        MsgBox Err.Description
                        Stop
                        Resume Exit_Help_TXT_GotFocus

                    End Sub

                    '//@COMBINATA_Help_TXT_GotFocus *** FINE ***
                    '//##################################################################################################################//

            03 GESTIRE L'EVENTO SU MODIFICA DELLA CASELLA COMBINATA PER CARICARE I MENU

               '//APRO COMANDO @?MENU@PAG.6 - @COMBINATA
                    '//##################################################################################################################//
                    '//Nota: questa routine chiama la funzione di gestione dei menu che si trova nel modulo : _
                            CASELLA_COMBINATA_Mdl_01_GestioneMENU
                    Private Sub COMANDO_s_TXT_AfterUpdate()

                    On Error GoTo Err_COMANDO_s_TXT_AfterUpdate

                        '//ATTIVO IL MENU INTERNO @MENU@HELP
                        '//imposto solo 3 parametri su 4 della funzione apri file
                       
                            Dim scelta_s As String
                            Dim Scelta_Pasth_s As String
                            Dim Scelta_MyFile_s As String
                            Dim Scelta_IDGestione_lng As Long

                        '//per evitare errore null
                        If Me.COMANDO_s_TXT > "" Then
                                scelta_s = Me.COMANDO_s_TXT
                        Else
                            '//se nullo esci dalla procedura per evitare errori
                            GoTo Exit_COMANDO_s_TXT_AfterUpdate
                        End If
                        
                        Select Case scelta_s
                            
                            
                            '//@MENU@PAGINA.6
                            Case "1_APRI_REPORT_FORM"
                            
                                    
                                    '//CHIAMA LA FUNZIONE GESTIONE MENU DELLA CASELLA COMBINATA, @FUNZIONE@MENU@CAELLA@COMBINATA
                                    '//---------------------------------------------------------------------------//
                                         '//
                                         pfFunction_GE_MENU_CASELLA_COMBINATA (scelta_s)
                                    '//---------------------------------------------------------------------------//
                                    
                                   
                             '//@APRO@I.PROGGETTI, @OGGETTI@MSYS@PROGETTI, @MENU@SYS@APRI@PROGETTI.TUTTI, @MENU@APRO@PROGETTI.TUTTI
                            Case "10_OGGETTI_MSys_ApriProgettiTutti"
                                    
                                    '//CHIAMA LA FUNZIONE GESTIONE MENU DELLA CASELLA COMBINATA, @FUNZIONE@MENU@CAELLA@COMBINATA
                                    '//---------------------------------------------------------------------------//
                                         '//
                                         pfFunction_GE_MENU_CASELLA_COMBINATA (scelta_s)
                                    '//---------------------------------------------------------------------------//
                                   
                           '//@HELP@APRI.PDF
                            Case "20_HELP_PROGETTI_ApriPdf"
                            
                                    '//CHIAMA LA FUNZIONE GESTIONE MENU DELLA CASELLA COMBINATA, @FUNZIONE@MENU@CAELLA@COMBINATA
                                    '//---------------------------------------------------------------------------//
                                         '//
                                         pfFunction_GE_MENU_CASELLA_COMBINATA (scelta_s)
                                    '//---------------------------------------------------------------------------//
                                
                          
                                
                            Case Else
                                MsgBox "Scelta non esistente: " & scelta_s, vbExclamation, "SCELTA HELP"
                                '//USCITA DALLA ROUTINE
                                GoTo Exit_COMANDO_s_TXT_AfterUpdate
                        End Select
                       


                    '//CHIUSURA E CONTROLLO ERRORI
                    '//-----------------------------------------------------------------------------------------//

                    Exit_COMANDO_s_TXT_AfterUpdate:
                        Exit Sub

                    Err_COMANDO_s_TXT_AfterUpdate:
                        MsgBox Err.Description
                        Stop
                        Resume Exit_COMANDO_s_TXT_AfterUpdate

                    End Sub

                    '//*** FINE ***
                    '//APRO COMANDO @?MENU@PAG.6
                    '//##################################################################################################################//







    FAQ
        @IMPOSTARE@MENU.CMB
        @CASELLA.COMBINATA@MENU
        @CREARE@MENU.CASELLA.COMBINATA, @CREARE.UN.MENU
        @COME@CREARE@UN.MENU

