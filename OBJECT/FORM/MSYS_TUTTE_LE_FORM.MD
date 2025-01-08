
FORM_MSYS
		Note
			Msys ha divers form per la  gestione degli oggetti:
			@Msys_CODICE---->@Oggetti@Msys
				ANAG_QryDF01_01_PARAMETRI_INDIV
					@query.estraggo.parametri.individuali
					@form.parametri.individuali


FORM DEL PROGETTO CARICATE
		Note
			Oggetto Msys che si riferisce al gruppo di form da caricare a seconda
			del progetto in gestione. Form basata sualla tabella oggetti MSys_TUTTI_PROG
			dove vengono salvati tutti i progetti. Le pagine di seguito indicate indicano
			le form che vengono inserite nel Tab Control.
			

FORM MASTER PRINCIPALE
	
	AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER,
		ridenominata al posto della MSys_FORM_M01_GESTIONE_OGGETTI_MSys	
		Note
			La form master principale di gestione di tutti i progetti
		Faq
			@FORM@MASTER@MSYS
			@MASTER@GESTIONE@PROGETTI.TUTTI

	    Tab_Control

			PAG_01
			PAG_02
			PAG_03
			PAG_04
			PAG_05
			PAG_06

			PAG_07 

				ESPORTA PROGETTI
				
				MSys_DF10_DATABASE_Frm01_M0_ESPORTA_CurrentProject
					Note
						FORM MASTER Query sottostante
						MSys_DF10_DATABASE_Qry01_01_SELECT_TUTTI_Current
							Query che estrae nella form master solo il progetto corrente con lo stesso
							codice progetto 

			 	MSys_DF10_DATABASE_Frm01_S01_ELENCO_Current
						
								@sottoform@esporta@progetto.corrente


					Note
						Form master di esportazione del progetto corrente
							 @FORM@ESPORTA.PROGETTO.CORRENTE@query
			                 @form@ESPORTA@PROGETTO.CORRENTE
					
					SottoForm
						MSys_DF10_DATABASE_Frm01_S01_ELENCO_CurrentProject				 								

			PAG_08
				..
			PAG_09
				...
			PAG_10

				VISUALIZZA I PROGETTI
					Nota
						la form che visualizza i progetti è la 
							Msys_DF11_PROGETTI_Frm00_M01_GESTIONE

									Questa è la form MASTER DELLA GESTIONE DEI PROGETTI cosi costituita da 2 sottoform

									01) Msys_DF11_PROGETTI_Frm00_M01_GESTIONE
									02) Msys_DF11_PROGETTI_Frm01_M01_Obj
										Query---> Msys_DF11_Qry01_01_PROGETTI
									03) Msys_DF11_PROGETTI_Frm02_S01_TIPO_OGGETTI
											Query---> Msys_DF11_PROGETTI_Frm02_S01_TIPO_OGGETTI

					struttura
						COD_PROG_s 
							Campo codice progetto con il doppio click apro la query che evidenzia tutte le definizioni dei progetti.
								@COD_PROG_s.EVENTO
								@EVENTO@dopioclick
								@VISUALIZZA@DEFINIZIONE@PROGETTI
								@DF11
								@DEFINIZIONE@PROGETTI
								@VISUALIZZO@PROGETTI



					FAQ
						@PROGETTI@GESTIONE
						@VISUALIZZA.I.PROGETTI
						@PAG_10
			
			PAG_11						 

			PAG_12

				MSys_TUTTI_PROG_FrmM01_01_PROGETTI
					
					Note
						Form che visualizza il progetto caricato
							@form@master@proggetti@caricati
							@proggetto@visualizzato
							@progetto@corrente
							@current.project.visualizzato
							@tab@12.progetti.caricati

					Modifiche
						2022.10.08
							Inserito nuovo campo  Note_PROGETTO_s, la form master progetti

					faq
						@FORM.PROGETTI


					CASELLA_COMBINATA_HELP

						Note
							Per la gestione dei progetti la l'analisi viene imposta nella casella combinata
							Help con due eventi il focus in cui viene impostata con i MENU per la gestione dei progetti e l'evento dopo aggiornamento in cui viene attivato il menu.
								FAQ
									@MENU@ANALISI@ANAGRAFICA, @MSYS@ANALISI@PROGETTI @MSYS@MENU
									@HELP@GESTIONE.PROGETTI; @HELP@PROGETTI
									@ANALISI@MENU@GESTIONE.PROGETTI,  @MENU@ANAGRAFICA,
									@EVENTI@COMBINATA@PROGETTI
									@MENU@PROGETTI.EVENTI 
									@1_ANALISI_PROGETTO_ANAGRAFICA



							EVENTI

								'//@COMBINATA_Help_TXT_GotFocus
								'//##################################################################################################################//
								Private Sub Help_TXT_GotFocus()

								On Error GoTo Err_Help_TXT_GotFocus




								    
								    '//IMPOSTO@LA@COMBINATA AL FOCUS DELLA CASELLA
								    '//------------------------------------------------------------------------------------------------------//
								    '//Note Se la proprietà RowSourceType è impostata a Elenco valori, _
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
								        Me.Help_TXT.RowSource = "1_ANALISI_PROGETTO_ANAGRAFICA;2_Mar;3_Mer;"
								        
								        
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

								'//APRO L'HELP  @?@HELP - @COMBINATA
								'//##################################################################################################################//
								Private Sub Help_TXT_AfterUpdate()

								On Error GoTo Err_Help_TXT_AfterUpdate

								    '//ATTIVO IL MENU INTERNO @MENU@HELP
								    '//imposto solo 3 parametri su 4 della funzione apri file
								   
								        Dim scelta_s As String
								        Dim Scelta_Pasth_s As String
								        Dim Scelta_MyFile_s As String
								        Dim Scelta_IDGestione_lng As Long



								    scelta_s = Me.Help_TXT
								    
								    Select Case scelta_s
								        '//@MENU@SCELTA.ANAGRAFICA
								        Case "1_ANALISI_PROGETTO_ANAGRAFICA"
								        
								                
								                '//CHIAMA LA FUNZIONE GESTIONE MENU DELLA CASELLA COMBINATA
								                '//---------------------------------------------------------------------------//
								                     '//
								                     pfFunction_GE_MENU_CASELLA_COMBINATA (scelta_s)
								                '//---------------------------------------------------------------------------//
								                
								               

								            
								        Case Else
								            MsgBox "Scelta non esistente: " & scelta_s, vbExclamation, "SCELTA HELP"
								            '//USCITA DALLA ROUTINE
								            GoTo Exit_Help_TXT_AfterUpdate
								    End Select
								   


								'//CHIUSURA E CONTROLLO ERRORI
								'//-----------------------------------------------------------------------------------------//
								Exit_Help_TXT_AfterUpdate:
								    Exit Sub

								Err_Help_TXT_AfterUpdate:
								    MsgBox Err.Description
								    Stop
								    Resume Exit_Help_TXT_AfterUpdate

								End Sub

								'//*** FINE ***
								'//APRO L'HELP  @?@HELP
								'//##################################################################################################################//
FORM GESTIONE PROGETTI
	NOTE
		Per attivare la form della gestione dei progetti occorre copiare dal database
		msys questi oggetti:
			LE SEGUENTI FORM
				MSys_DF11_}----------------------------------------------------@
				Msys_DF11_PROGETTI_Frm00_M01_GESTIONE
				Msys_DF11_PROGETTI_Frm01_M01_Obj
				Msys_DF11_PROGETTI_Frm02_S01_TIPO_OGGETTI


			LE SEGUENTI QUERY
				Msys_DF11_PROGETTI_Qry00_}-------------------------------------@
				Msys_DF11_PROGETTI_Qry01_10_GROUP_TUTTI
				Msys_DF11_Qry01_01_PROGETTI
				Msys_DF12_Qry01_}----------------------------------------------@
				Msys_DF12_Qry01_01_TIPO_OGGETTI
				MSys_FORMS_Qry01_71_INSERT_toMsys_DF12
				MSys_MACRO_Qry01_71_INSERT_toMsys_DF12
				MSys_MODULI_Qry01_71_INSERT_toMsys_DF12
				MSys_QUERY_Qry01_71_INSERT_toMsys_DF12
				MSys_REPORT_Qry01_71_INSERT_toMsys_DF12
				MSys_TABELLE_Qry01_71_INSERT_toMsys_DF12



UTILITA

	ELENCO OGGETTI ANAGRAFICA CONDOMINIALE DA ESPORTARE
			Note
				Attenzione per utilizzare questi oggetti occorre riportare la tabelle su execel e poi copiarle e ricaricarle
				nel db di access.

	OGGETTI TABELLE NECESSARIE PER IL FUNZIONAMENTO

			elenco xls

			
								COD_PROGETTO_s	TIPOGGETTO_s	NRO_OGGETTO	Note	Scel_b	>>Name1	>>Attribute	NROGGETTO	LvExtra	Expression	Name2	EVENTO	ATTIVITAEVENTO	METODO	MODFORM	CODATTIVITA	DATAINS	DATAAGG	TIMEOPER	IDOGGETTO	COMANDO
				CDM_ANGRAFE	TABLE	1	Per le tabelle anagrafiche	0	ANAG_{@===================================================@}_Tab											05/06/2022		19:58:32	2778	
				CDM_ANGRAFE	TABLE	2	Per le tabelle anagrafiche	0	ANAG_DF01_PARAMETRI_INDIV											05/06/2022		19:58:32	2779	
				CDM_ANGRAFE	TABLE	3	Per le tabelle anagrafiche	0	ANAG_TB01_PROPRIETARI											05/06/2022		19:58:32	2780	
				CDM_ANGRAFE	TABLE	4	Per le tabelle anagrafiche	0	ANAG_TB01_PROPRIETARI_TMP											05/06/2022		19:58:32	2781	
				CDM_ANGRAFE	TABLE	5	Per le tabelle anagrafiche	0	ANAG_TB02_GESTIONE_PROPRIETARI											05/06/2022		19:58:32	2782	
				CDM_ANGRAFE	TABLE	6	Per le tabelle anagrafiche	0	ANAG_TB03_PROPRIETA_CONDOMINIALI											05/06/2022		19:58:32	2783	
				CDM_ANGRAFE	TABLE	7	Per le tabelle anagrafiche	0	ANAG_TB70_ASSEMBLEA											05/06/2022		19:58:32	2784	

	OGGETTI QUERY NECESSARIE PER IL DI FUNZIONAMENTO

			elenco xls


					COD_PROGETTO_s	TIPOGGETTO_s	Attribute	COMANDO	NRO_OGGETTO	LvExtra	Expression	Note	Scel_b	>>Name1	Name2	EVENTO	ATTIVITAEVENTO	METODO	MODFORM	CODATTIVITA	DATAINS	DATAAGG	TIMEOPER	IDOGGETTO
					CDM_ANGRAFE	QUERY			1				FALSO	ANAG_Qry00_{@=================================================@}							02/07/2022		9:13:48 PM	1436
					CDM_ANGRAFE	QUERY			2				FALSO	ANAG_QryDF01_00_PARAMETRI_INDIV_}-------------------------SELECT							02/07/2022		9:13:48 PM	1437
					CDM_ANGRAFE	QUERY			3				FALSO	ANAG_QryDF01_01_PARAMETRI_INDIV							02/07/2022		9:13:48 PM	1438
					CDM_ANGRAFE	QUERY			4				FALSO	ANAG_QryDF01_60_PARAMETRI_INDIV_}-------------------------UPDATE							02/07/2022		9:13:48 PM	1439
					CDM_ANGRAFE	QUERY			5				FALSO	ANAG_QryDF01_61_PARAMETRI_INDIV_Aggiorno_CodVersamenti							02/07/2022		9:13:48 PM	1440
					CDM_ANGRAFE	QUERY			6				FALSO	ANAG_QUERY_Qry01_00_}------------------------------------------@							02/07/2022		9:13:48 PM	1441
					CDM_ANGRAFE	QUERY			7				FALSO	ANAG_MsysQry01_01_SELECT_OggettiFORM							02/07/2022		9:13:48 PM	1442
					CDM_ANGRAFE	QUERY			8				FALSO	ANAG_TABELLE_Qry00_00_}-----------------------------------SELECT							02/07/2022		9:13:48 PM	1443
					CDM_ANGRAFE	QUERY			9				FALSO	ANAG_TABELLE_Qry01_01_SELECT_TUTTE							02/07/2022		9:13:48 PM	1444

	OGGETTI FORM NECESSARIE PER LA GESTIONE DEL PROGETTO ANAGRAFE
			elenco xls

				NRO_OGGETTO	TIPOGGETTO	NOMEOGGETTO	NOMEpagina	VISIBLE_b	>>NOTA OGGETTO	COD_PROGETTO	DATAINS_d	DATAAGG_d
				1	SottoForm_01	ANAG_TABELLE_Frm01_S01_TUTTE	01 TABELLE MSys	-1	Pagina 1	CDM_ANGRAFE	10/06/2022	
				3	SottoForm_03	ANAG_MsysQry01_01_SELECT_OggettiQUERY	02 QUERY oggetti Anagrafica	-1	Pagina 3	CDM_ANGRAFE	10/06/2022	
				5	SottoForm_05	ANAG_TB01_PROPRIETARI_Frm01_S01_TUTTI	05 ANAGRAFE PROPRIETARI	-1	Pagina 5	CDM_ANGRAFE	10/06/2022	
				7	SottoForm_07	MSys_DF10_DATABASE_Frm01_M0_ESPORTA	07 ESPORTA OGGETTI	-1	Pagina 7	CDM_ANGRAFE	10/06/2022	31/07/2022
				8	SottoForm_08	ANAG_FrmDF01_S01_PARAMETRI_INDIV	08 PARAMETRI INDIVIDUALI	-1	Pagina 8	CDM_ANGRAFE	10/06/2022	31/07/2022
				10	SottoForm_10	Msys_DF11_PROGETTI_Frm00_M01_GESTIONE	10 PROGETTI	-1	Pagina 10	CDM_ANGRAFE	10/06/2022	
				11	SottoForm_11	MSys_SERVIZI_Frm01_S01_TUTTI	11 TUTTI I SERVIZI+ESPORTA	-1	Pagina 11	CDM_ANGRAFE	10/06/2022	
				12	SottoForm_12	MSys_TUTTI_PROG_FrmM01_01_PROGETTI	12 FORM DEL PROGETTO CARICATE	-1	Pagina 12	CDM_ANGRAFE	10/06/2022	