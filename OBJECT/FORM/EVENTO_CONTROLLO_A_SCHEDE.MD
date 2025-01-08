EVENTO_CONTROLLO_A_SCHEDE.MD


	Note
		@tab@controll_(evento del controllo a schede)

		Per l'attivazione degli eventi del controllo a schede al cambio di
		pagina deve essere utilizzata la seguente routine:






	ROUTINE_EVENTO_CHANGE
		
		'//*********************************************************//
			'//
			'//                   @EVENTI@FORM
			'//
			'//*********************************************************//


			'//EVENTO CHANGE DEL CONTROLLO A SCHEDE
			'//=========================================================================================//
			'//@EVENTO@TABE@CONTROLL_(il @cambio di @pagine)
			'//note:    ATTENZIONE il controllo a schede si deve chiamare TabCtl_01 _
			            le pagine : pag_01 ... e le _
			            sottoform : SottoForm_01 ...
			Private Sub TabCtl_01_Change()

			'//i comandi gli argomento e le pagine correlate *** da utilizzare **
			    Dim ixCommand As Integer
			    Dim vxArgument As Variant
			    Dim sxPagCorrente_Caption As String
			    
			                 

			    On Error GoTo Err_TabCtl_01_Change
			        
			       '//Riaggiorno la form principale
			       'Me.Recalc
			       'Me.Requery
			       
			    'PAGINE
			    '____________________________________________________________________________________________________________
			        
			        
			        'RECUPERO IL NOME PAGINA CORRENTE
			        '............................................................................
			            'reset variabili
			            
			            ixPage = Me.TabCtl_01
			            
			            'chiamo la funzione
			            'sxPagCorrente_Caption = pfunc_Recupero_NOME_PAGINA_CORRENTE(Dim sxFORM_CHIAMANTE_GENERALE_CORRENTE, 1)
			        '............................................................................
			        
			        
			        If Me.TabCtl_01.Value = 0 Then
			               
			                    'PG_01
			                    '............................................................................
			                    'Imposto il valore della casella combinata
			                    'valore = 0 : pannello lotto; valore = 1 : ambo ecc.
			                    'Me.Cmb_01.Value = 1
			                        
			                            'Imposta i parametri del comando per l'ambo
			                            ixCommand = 1
			                            vxArgument = 412
			                            'la scheda corrente attiva
			                            sxPagCorrente_Caption = ""
			                            
			                          
			                         'Me.SottForm_Pg_02.Requery
			                    
			                        'Aggiorno Sottoform_01
			                        'Me.SottForm_Pg_01.Requery
			                '............................................................................
			    'Pag 2
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 1 Then
			    
			                    
			                    'PG_02
			                    '............................................................................
			                    'Imposto il valore della casella combinata
			                    
			                            'Imposta i parametri del comando per l'ambo
			                            ixCommand = 1
			                            vxArgument = 413
			                            'la scheda corrente attiva
			                            'sxPagCorrente_Caption = "TERNO"
			                '............................................................................
			                
			                        '//REQUERY DELLA SOTTOFORM_2
			                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
			                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
			                        'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_02].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco_ANNUALITA_TMP_Folium"
			                        
			                         Me.SottoForm_02.Requery
			                
			    'Pag 3
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 2 Then
			    
			                
			                'PG_03 - TAB CONTROLL -> TabCtl_01.value=2
			                '............................................................................
			                    'Imposto il valore della casella combinata
			                        'Imposta i parametri del comando per l'ambo
			                            ixCommand = 1
			                            vxArgument = 414
			                            'la scheda corrente attiva
			                            'sxPagCorrente_Caption = "QUADERNA"
			                        
			                        '............................................................................
			                
			                        '//REQUERY DELLA SOTTOFORM_2
			                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
			                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
			                        Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_03].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco_ANNUALITA_TMP_Folium"
			                        
			                        
			                         'Me.SottForm_Pg_03.Requery
			                        
			                '............................................................................
			                    
			                
			    'Pag 4
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 3 Then
			    
			                
			                'PG_04
			                '............................................................................
			                    'Imposto il valore della casella combinata
			                    
			                        
			                            'Imposta i parametri del comando per l'ambo
			                            ixCommand = 1
			                            vxArgument = 415
			                            'la scheda corrente attiva
			                            sxPagCorrente_Caption = "?"
			                               
			                        '............................................................................
			                
			                        '//REQUERY DELLA SOTTOFORM_4
			                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
			                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
			                        Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_04].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_DET"
			                        
			                        'Me.SottForm_Pg_04.Requery
			                        
			                '............................................................................
			            
			    '____________________________________________________________________________________________________________

			                 
			    'Pag 5
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 4 Then
			    
			                'Aggiorno form
			                'Me.Pg_05.Requery
			                
			                
			                
			                'PG_05 TAB CONTROLL -> TabCtl_01.value=4
			                '............................................................................
			                    'Imposto il valore della casella combinata
			                   '............................................................................
			                
			                        '//REQUERY DELLA SOTTOFORM_3
			                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
			                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
			                        Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_05].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco_ANNUALITA_TMP_DGC"
			                        
			                  
			                    
			                '............................................................................
			                
			                '//REQUERY DELLA SOTTOFORM_5
			                'Me.SottForm_Pg_05.Requery
			                         
			    '____________________________________________________________________________________________________________

			    'Pag 6
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 5 Then
			    
			                'Aggiorno form
			                'Me.Pg_06.Requery
			                
			                
			                
			                'PG_06 TAB CONTROLL -> TabCtl_01.value=5
			                '............................................................................
			                        '//REQUERY DELLA SOTTOFORM_2
			                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
			                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
			                        Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_06].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco_ANNUALITA_TMP_DCC"
			                        
			                         
			                        
			                '............................................................................
			                
			                '//REQUERY DELLA SOTTOFORM_6
			                'Me.SottForm_Pg_06.Requery
			                         
			    '____________________________________________________________________________________________________________
			        
			            
			            
			    'Pag 7
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 6 Then
			    
			                'Aggiorno form
			                'Me.Pg_07.Requery
			                
			                
			                
			                'PG_07 ???
			                '............................................................................
			                    'Imposto il valore della casella combinata
			                      '//REQUERY DELLA SOTTOFORM_2
			                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
			                        '//MascheraPrincipale--------->Sottoform02..............................Sottoform03...........................>RecordoSource......="NomeQuery"
			                         'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco"
			                         'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco"
			                          'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco"
			                          '[Forms]![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07].[Form]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].[Form].[RecordSource] = "LLPP_ATTI_Qry01-12_Gestione_Elenco"
			                          'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA].[LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
			                          'Application.Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA].[LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
			                          'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
			                          
			                          'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
			                         'Application.Forms.Item(14).Form
			                          'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
			                            Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"

			                        'DB_AGG.Form_AA_GEST_FrmM00_GESTIONE_MASTER.SottForm_Pg_07![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
			                '............................................................................
			                
			                '//REQUERY DELLA SOTTOFORM_7
			                'Me.SottForm_Pg_07.Requery
			                         
			    '____________________________________________________________________________________________________________
			            
			        

			            
			    'Pag 8
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 7 Then
			    
			                'Aggiorno form
			                'Me.Pg_08.Requery
			                
			                
			                
			                'PG_08
			                '............................................................................
			                    'Imposto il valore della casella combinata
			                    '//REQUERY DELLA SOTTOFORM_2
			                     '//Cambio proprieta record source della sottoform caricata e il requery è automatico
			                     '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
			                     'Forms![AA_GEST_FrmM00_GESTIONE_MASTER]![SottForm_Pg_07].Form.RecordSource = ""
			                                                                                       
			                    
			                        
			                '............................................................................
			                
			                '//REQUERY DELLA SOTTOFORM_8
			                'Me.SottForm_Pg_08.Requery
			                         
			    '____________________________________________________________________________________________________________
			            
			                
			                
			                
			            
			    'Pag 9
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 8 Then
			    
			                'Aggiorno form
			                'Me.Pg_09.Requery
			                
			                
			                
			                'PG_09
			                '............................................................................
			                    'Imposto il valore della casella combinata
			                    
			                        
			                '............................................................................
			                
			                '//REQUERY DELLA SOTTOFORM_9
			                'Me.SottForm_Pg_09.Requery
			                
			                         
			    '____________________________________________________________________________________________________________
			            


			            
			    'Pag 10
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 9 Then
			    
			                'Aggiorno form
			                'Me.Pg_10.Requery
			                
			                
			                
			                'PG_10
			                '............................................................................
			                    'Imposto il valore della casella combinata
			                    
			                        
			                '............................................................................
			                
			                '//REQUERY DELLA SOTTOFORM_10
			                'Me.SottForm_Pg_10.Requery
			                         
			    '____________________________________________________________________________________________________________
			            

			            
			    'Pag 11
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 10 Then
			    
			              
			    '____________________________________________________________________________________________________________
			            

			    'Pag 12
			    '____________________________________________________________________________________________________________
			        ElseIf Me.TabCtl_01.Value = 11 Then
			    
			              
			    '____________________________________________________________________________________________________________
			            
			   
			                
			        
			    End If
			    
			    
			    '//Aggiorno la form master
			    'Me.Recalc
			    'Me.Requery
			                

			'USCITA  E GESTIONE ERRORI
			'..............................................................................................................

			Exit_TabCtl_01_Change:
			    Exit Sub

			Err_TabCtl_01_Change:
			    MsgBox Err.Description
			    
			    Resume Exit_TabCtl_01_Change


			End Sub


			'//EVENTO CHANGE DEL CONTROLLO A SCHEDE
			'//=========================================================================================//
