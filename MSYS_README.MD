
FAQ
	==========================================================================================
										LE FAQ
	==========================================================================================
			TABELLE PER L'ESPORTAZIONE DEGLI OGGETTI
				per esportare gli oggetti di un progetto vengono utilizzate le seguenti tabelle:

					Msys_ESPORTA_DB_EST 			= questa tabella contiene la path ed il database
					Msys_ESPORTA_ELENCO_OGGETTI_DB 	= questa tabella contiene l'elenco degli oggetti come
													query, form, report ecc. da esportare nella path della
													precedente tabella.

			TABELLE PER LA COSTRUZIONE DI UN PROGETTO
				Per la costruzione di un progetto e legare gli oggetti ad un codice vengono utilizzate
				le seguenti tabelle:
					MSys_DF02_TIPO_OGGETTO				= questa tabella definisce gli oggetti di un progetto
														che sono 5, Table, form, Report, Query e Macro

					Msys_DF11_PROGETTI					= la tabella df11, definisce il progetto ed un codice
														chiave e la path di dove si trova il file di gestione
														che di solito è GE_PROGETTO

					quindi vengono utilizzate le tabelle DF02 E DF11 PER LA DEFINIZIONE.



MSYS_REDME

	PROGETTO_MSYS_REDME
				Note
					il progetto MySys si riferisce al progetto relativo a tutti gli oggetti Access
					utili per gli altri progetti.
					Il progetto Msys si trova in questa path e file:
						c:\CASA\CDM\GE_CDM\MSYS\MSYS_MDB\MSYS_N01_TABELLE.mdb
						@PROGETTO@MSYS



					ATTENZIONE PATH DEL PROGETTO IN UFFICIO è QUESTA:
						c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\


				DB
					MSYS_N01_TABELLE.mdb
						si trova in questa path:
					c:\CASA\CDM\GE_CDM\MSYS\MSYS_MDB\



				FORM MASTER
					MSys_FORM_M01_GESTIONE_OGGETTI_MSys
						Questa è la form per la gestione di tutti gli oggetti msys e viene utilizzata
						per gli altri progetti.
						Permette di caricare i vari oggetti necessari al progetto in questione;

							Per tutti i dettagli fare riferimento a questo codice: @FORM@MASTER@MSYS

							Attenzione se cambi il nome alla form master questa procedura garantisce
							il suo funzionamento:
								@cambio@denominazione@form@master_(se cambi denominazione alla form master viene utilizzata ugualmente)

							Attenzione 


							'//@ERRORE@RECORD@VUOTO, se esiste il progetto da caricare, MA NON HA FORM DA IMPLEMENTARE allora il msg di errore _
	                                            + uscita dalla routine,
	                                        DoCmd.OpenQuery "MSys_TUTTI_PROG_SELECT_Qry01_01_PROGETTI", acViewNormal
	                                        

				GESTIONE_PROGETTO
					TODO: provo ad eliminare tutti gli oggetti senza filtro in db mdb esterno

					

					TODO
					sono da fare i seguenti miglioramenti per il progetto:
						1) Eliminare tutte le form  che non hanno il codice MSys_ perche appartengono 
							agli altri progetti. 
								DONE
									2023.02.11 FORM ELIMINATE = sono state eliminate tutte le form che 
												non appartegono al progetto Mys

						2) Eliminare tutte le QUERY  che non hanno il codice MSys_ perche appartengono 
							agli altri progetti.

								DONE
									2023.02.11 QUERY ELIMINATE = sono state eliminate tutte le QUERY che 
												non appartegono al progetto Mys

						3) Eliminare tutti I MODULI  he non hanno il codice MSys_ perche appartengono 
							agli altri progetti.
		


		FORM_CARICAMENTO_AUTOMATICO
			Il progetto FORM_M01_MASTER_base_(CaricaSottoform).mdb contiene la form MASTER per
			il caricamento automatico delle sottoform in numero max 12

		FORM
			MSys_FORM_M01_}------------------------------------------------@
			MSys_FORM_M01_GESTIONE_OGGETTI_MSys
			AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER
				Note
					form a caricamento automatico di 12 pagine
					Le due form sono le stesse solo la master dovrebbe essere quella che puo essere utilizzata
					per tutti i progetti, e quindi potrebbe rimanere mentre la GESTIONE_OGGETTI_Msys verrebbe
					cancellata appena finito la stesura del progetto.




						@form@caricamento@automatico_(La form @MASTER per caricare automaticamente @12@pagine)



				FUNZIONAMENTO
					01)FORM -> EVENTO @FORM@LOAD@PAGINE@TUTTE
						All'apertura della form carico tutto le pagine del tab controll per un  max di 12

						@RESET@PAGINE_(sempre all'apertura vengono resettare le sottoform le pagine del tab controll)

						@ATTIVO@FUNZIONE@CARICA@PAGINE_(viene attivata la funzione che carica le pagine delle sottoform)

						@FUNZIONE@CARICA@PAGINE@TAB@CONTROLL_(eseguo un ciclo sulla tabella MSys_TUTTI_PROG per trovare il progetto selezionato a True da caricare)


						@FUNZIONE@CARICA@PAGINE@TAB@CONTROLL_(recupero i parametri CODICE PROGETTO TIPO OGGETTO NOME OGGETTO NOME PAGINA VIBILE e chiamo la funzione imposta le sottoform)

						@CHIAMO@LA@ROUTIN@IMPOSTA@LE@SOTTOFORM_(per impostare le sottoform e passo 4 parametri)


						@ROUTINE@IMPOSTA@LE@SINGOLE@SOTTOFORM_(con la ROUTINE imposta le sottoform MA attivo la classe ricerca oggetto nel db)



						@ROUTINE@IMPOSTA@LE@SINGOLE@SOTTOFORM_(ma prima controllo il Parametro Vible = True e poi attivo la classe di controllo esistenza sottoform, infine solo se la form esiste la carico nella pagina scelta)




						@faq@come@caricare_(Le@pagine delle 12 sottoform sono a @caricamento@automatico
							)


					02) @PROGETTO@NULLO
						Nel caso che nessun progetto viene selezionato allora viene emesso un avviso di 
						progetto nullo, vedi codice
							@PROGETTO@NON@SELEZIONATO@ ...


					03) COMANDI O MENU
						Nella form MASTER, i menu devono cambiare nel cambio del progetto.
						Per ottenere questo risultato occorre quindi:
						a) creare una tabella tmp dei comandi denominata 	
								MSys_DF01_COMANDI_TMP
							In questa tabella devono essere salvati i comandi temporanei che attengono al progetto e poi caricati e visualizzati nella form master.
							Prima modifico la tabella originale e tmp con il codice progetto inserendo questo
							campo.
								COD_PROGETTO_s
							HO anche modificato le chiavi delle tabelle comandi e comandi tmp con questi due
							campi:
								COMANDO_s
								COD_PROGETTO_s
								PROGETTO_s

							- Ho creato la query di RAGGRUPPAMENTO per visualizzare i comandi assegnati ai
								progetti:
									MSys_DF01_COMANDI_Qry01_01_GROUP_TUTTI
									codice: comandi

							- Ho creato la query di accodamento dei comandi tmp in base al codice progetto.
									codice: comandi
								


						B)
							@2024_03_11_CREATO_INSERIMENTO_CODICI_PROGETTO_IN_TMP

							ho creato anche la tabella tmp relativo a i PROGETTI tmp dove vengono salvati
							solo 
								MSys_TUTTI_PROG_TMP

							Ho aggiunto e creato una query di accodamento del progetto scelto denominata
								MSys_TUTTI_PROG_INSERT_Qry01_01_PROGETTI_in_TMP

							questa query di accodamento è un base per l'operazione di cancellazione del progetto nella tabella
							tmp ed il successivo accodamento nella tabella
								MSys_TUTTI_PROG_TMP
							dei 3 codice del progetto scelto e questo è un esempio:
								NRO_OGGETTO_i	COD_PROGETTO_s				DENOM_PROGETTO_s
								5				FORM_AUTOCARICA	PROGETTO 	FORM AUTOCARICANTE


							vedi codice :
								@INSERTI@CODICI@PROGETTO@TMP   ...


						C)
							dopo aver inserito i codice del progetto tmp occorre creare:

								Cmb.01 = MODIFICO IL NOME DELLA COMBINATA 01 IN Cmb_01_TXT
									Ho cambiato la path dei comandi ....

									@cmbd_01.casella.combinata.evento_(APRO GRUPPO COMANDI CASELLA COMBINATA 01)
								TODO: fare una query che estrae i menu con il codice progetto tmp

								TODO: aggiornare il campo Cmb_01
									   CON IL MENU  del codice selezioto ...



	   REGISTRAZIONE_DELLE_LIBRERIE_DLL_(cmd Dos)

	   		Per caricare  la libreria OCX  che si trova qui:
		   	c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\DLL_LIBRERIE\MSCOMCTL.OCX

		   			@DLL@CMD@DOS_(comando di apertura della shell del @registro@di@sistema)
		   			@REGISTRO@DLL@CMD@DOS_(comando DOS PER LA REGISTRAZIONE DELLE LIBRERIE)
		   		eseguire il comando
				regSvr32 con il file .ocx

				per registrare manualmente la ddl ocorre 
				1) aprire il cmd come amministratore 
				2) cd.. radice posizionando in 
					c:\Windows\SysWOW64\
				3) eseguire il comando:
					regSvr32 MSCOMCTL.OCX


	GIORNALE_OPERAZIONI


			start "apri GIORNALE DELLE OPERAZIONI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\TUTORIAL\APRI_GIORNALE_OPERAZIONI_PROGETTO_MSYS.bat"

		2024 RACCORDO tra la tabella Project e la tabella quote. 
	PROGETTI_Msys_TB01_PROJECT
		2025 CORREZIONI
			RACCORDO tra la tabella ELENCO objProgetti dove vengono salvati gli oggetti per ogni 
			progetto e la tabella QUOTE, deputata alla costruzione dei codici per ogni progetto. 
				PROGETTI_Msys_TB03_ELENCO_ObjProgetti
				QUOTE_MsysTB01_CODICI_QUOTE
				J5KZj8s34pSCRtagblJ60vjS




			


                                



