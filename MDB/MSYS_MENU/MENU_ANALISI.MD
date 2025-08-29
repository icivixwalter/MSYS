MENU_ANALISI.MD

	Note
		@analisi@menu_(analisi sui menu)

		Per i menu vi il pogetto MENU_GESTIONE.mdb	  che è duputato
		alla gestione dei menu in collegamento con le altre due 
		voci mdb
			MENU_TB01_VOCI.mdb			= qui vengo salvate le voci dei menu
			MENU_TB02_FUNZIONI.mdb		= qui sono trascritti le funzioni dei
										singoli menu.


			MENU_TB03_OGGETTI.mdb		= quest'ultimo progetto è da utilizzare??


			La path della gestione dei menu:
				@apri@path@menu_(@path@menu si trova in questa directory detta
					anche @menu@msys o @menu@mysys)

				"C:\Program Files\totalcmd\TOTALCMD64.EXE" left: c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU\, right: c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\ 



	VOCI MENU
		Note
			la tabella che contiene le voci menu è la seguente:

		GEST_MENU_TB01_VOCI
			Note
				vengono salvati i menu principale e la tabella si collega con
				la successiva delle funzioni tramite
				DATABASE=C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\MENU\MENU_TB01_VOCI.mdb;TABLE=GEST_MENU_Tb01_VOCI

				campi master: CodMenu_s	
				campi client: CodMenu_s	

				CODICE MENU VOCI ----> GEST_MENU_TB01
			QUERY
				GEST_MENU_Tb01Qry01_{@===================================@}
				GEST_MENU_Tb01Qry01_00_}-------------------------------VOCI_MENU
				GEST_MENU_Tb01Qry01_01_VOCI
					Note
						@estraggo_(tutti i @menu, schema )
					schema
						
						SELECT 
							GEST_MENU_Tb01_VOCI.CodMenu_s, GEST_MENU_Tb01_VOCI.Menu_s, GEST_MENU_Tb01_VOCI.NoteMenu_s, GEST_MENU_Tb01_VOCI.COMANDO_s, GEST_MENU_Tb01_VOCI.COD_PROGETTO_s, GEST_MENU_Tb01_VOCI.PROGETTO_s, GEST_MENU_Tb01_VOCI.bCanc, GEST_MENU_Tb01_VOCI.bAgg, GEST_MENU_Tb01_VOCI.bScel, GEST_MENU_Tb01_VOCI.IDKeyArch, GEST_MENU_Tb01_VOCI.IdOrd, GEST_MENU_Tb01_VOCI.Msg01, GEST_MENU_Tb01_VOCI.ID, GEST_MENU_Tb01_VOCI.DATA_INS, GEST_MENU_Tb01_VOCI.DATA_AGG, GEST_MENU_Tb01_VOCI.TIMEOPER

						FROM 
							GEST_MENU_Tb01_VOCI
						WITH 
							OWNERACCESS OPTION;

				GEST_MENU_Tb01Qry01_01_VOCI_TMP
					Note
						@estraggo_(salvo il menu selezionato im tmp, @menu@tmp )
					schema
						SELECT 
							GEST_MENU_Tb01_VOCI_TMP.CodMenu_s, GEST_MENU_Tb01_VOCI_TMP.Menu_s, GEST_MENU_Tb01_VOCI_TMP.NoteMenu_s, GEST_MENU_Tb01_VOCI_TMP.COMANDO_s, GEST_MENU_Tb01_VOCI_TMP.COD_PROGETTO_s, GEST_MENU_Tb01_VOCI_TMP.PROGETTO_s, GEST_MENU_Tb01_VOCI_TMP.bCanc, GEST_MENU_Tb01_VOCI_TMP.bAgg, GEST_MENU_Tb01_VOCI_TMP.bScel, GEST_MENU_Tb01_VOCI_TMP.IDKeyArch, GEST_MENU_Tb01_VOCI_TMP.IdOrd, GEST_MENU_Tb01_VOCI_TMP.Msg01, 
							GEST_MENU_Tb01_VOCI_TMP.ID, GEST_MENU_Tb01_VOCI_TMP.DATA_INS, GEST_MENU_Tb01_VOCI_TMP.DATA_AGG, GEST_MENU_Tb01_VOCI_TMP.TIMEOPER

							FROM 
								GEST_MENU_Tb01_VOCI_TMP
							WITH 
								OWNERACCESS OPTION;

				GEST_MENU_Tb01Qry01_01_VOCI_TMP_Delete
					Note
						cancello i dati nella tabella delle VOCI TMP

					Schema
						DELETE 
							GEST_MENU_Tb01_VOCI_TMP.*
							FROM 
							GEST_MENU_Tb01_VOCI_TMP
							WITH OWNERACCESS OPTION;

					


			FORM
				GEST_MENU_Tb01Frm00_{@======================================@}
				GEST_MENU_Tb01Frm01_}------------------------------------------@
				GEST_MENU_Tb01Frm01_M01_GE_VOCI
				GEST_MENU_Tb01Frm01_S01_VOCI
					Proprieta
						Master/Detail
							form per le voci di menu collegata alla sottofom
							tb01_FUNZIONI con queste chiavi

								campi master: CodMenu_s	
								campi client: CodMenu_s	

					CAMPI
						CodMenu_s_txt_DblClick
							nel campo codice menu attivo il salvataggio
							dei codici menu e condominio
							
							schema
					
								sSql = ""
								sSql = sSql & ""
								sSql = sSql & "INSERT INTO "
								sSql = sSql & "GEST_MENU_Tb01_VOCI_TMP "
								sSql = sSql & "( CodMenu_s, COD_PROGETTO_s ) "
								sSql = sSql & "SELECT "GeCondom" AS codice, "
								sSql = sSql & "GeCasa" AS CODICE_PROGETTO "
								sSql = sSql & "WITH OWNERACCESS OPTION;"



				GEST_MENU_Tb01Frm01_S01_VOCI_TMP
					Note
						form che contiene i menu tmp quando vengono selezionati
						e salvati nella tabella voci tmp

				GEST_MENU_Tb02Frm00_{@======================================@}
				GEST_MENU_Tb02Frm01_}------------------------------------------@
				GEST_MENU_Tb02Frm01_M01_GE_FUNZIONI
				GEST_MENU_Tb02Frm01_S01_FUNZIONI
					Note
						l'archivio delle funzioni

					EVENTI
						CodMenu_s_txt_DblClick = SALVO MENU IN TMP
						'//NOTE:  con il doppio click salvo il menu nella tabella tmp delle voci menu GEST_MENU_Tb01_VOCI_TMP
				GEST_MENU_Tb02Frm01_S01_FUNZIONI_TMP
					Note
						le funzioni TEMPORANEE salvate nella tabella delle FUNZIONI TMP
				UTILITA_MsysDbEst_Frm12_}--------------------------------------@
				UTILITA_MsysDbEst_Frm12_SERVIZI_TUTTI
					Note 
						utilita per i servizi tutti compresa l'importazione degli
						oggetti.
						Per ora imposta solo le form.
							todo: sistemare le altre importazioni.




		GEST_MENU_TB02_FUNZIONI
			Note
				tutte le funzioni collegate al menu principale

				CODICE VOCI FUNZIONI ----> GEST_MENU_TB02

			QUERY
				GEST_MENU_Tb02Qry01_}----------------------------------------@
				GEST_MENU_TB02Qry01_01_FUNZIONI
					Note
						@funzioni_(estraggo le funzioni collegate ai menu)
					schema

						SELECT 
								GEST_MENU_TB02_FUNZIONI.CodFunzioni_s, 
								GEST_MENU_TB02_FUNZIONI.Funzioni_s, 
								GEST_MENU_TB02_FUNZIONI.ClassFunzioni_s, GEST_MENU_TB02_FUNZIONI.NoteFunzioni_s, 
								GEST_MENU_TB02_FUNZIONI.CodMenu_s, 
								GEST_MENU_TB02_FUNZIONI.bCanc, 
								GEST_MENU_TB02_FUNZIONI.bAgg, 
								GEST_MENU_TB02_FUNZIONI.bScel, 
								GEST_MENU_TB02_FUNZIONI.IDKeyArch, 
								GEST_MENU_TB02_FUNZIONI.IdOrd, 
								GEST_MENU_TB02_FUNZIONI.Msg01, 
								GEST_MENU_TB02_FUNZIONI.DATA_INS, 
								GEST_MENU_TB02_FUNZIONI.DATA_AGG, 
								GEST_MENU_TB02_FUNZIONI.TIMEOPER, 
								GEST_MENU_TB02_FUNZIONI.ID, "->" 	  	AS [CampiTxt],
								GEST_MENU_TB02_FUNZIONI.CodFunzioni_s 	AS [CODICE_FUNZIONI],
								GEST_MENU_TB02_FUNZIONI.Funzioni_s 		AS [FUNZIONI],
								"chiavi" 								AS [TxtKey],
								GEST_MENU_TB02_FUNZIONI.CodMenu_s 		AS [KeyCodMenu_s]

							FROM 
								GEST_MENU_TB02_FUNZIONI
								ORDER BY GEST_MENU_TB02_FUNZIONI.ID
							;

				GEST_MENU_TB02Qry01_01_FUNZIONI_TMP
					Note
						le funzioni salvate nella tabella FUNZIONI TMP
					schema
						SELECT 
								GEST_MENU_TB02_FUNZIONI_TMP.CodFunzioni_s, 
								GEST_MENU_TB02_FUNZIONI_TMP.Funzioni_s, 
								GEST_MENU_TB02_FUNZIONI_TMP.ClassFunzioni_s, GEST_MENU_TB02_FUNZIONI_TMP.NoteFunzioni_s, 
								GEST_MENU_TB02_FUNZIONI_TMP.CodMenu_s, 
								GEST_MENU_TB02_FUNZIONI_TMP.bCanc, 
								GEST_MENU_TB02_FUNZIONI_TMP.bAgg, 
								GEST_MENU_TB02_FUNZIONI_TMP.bScel, 
								GEST_MENU_TB02_FUNZIONI_TMP.IDKeyArch, 
								GEST_MENU_TB02_FUNZIONI_TMP.IdOrd, 
								GEST_MENU_TB02_FUNZIONI_TMP.Msg01, 
								GEST_MENU_TB02_FUNZIONI_TMP.DATA_INS, 
								GEST_MENU_TB02_FUNZIONI_TMP.DATA_AGG, 
								GEST_MENU_TB02_FUNZIONI_TMP.TIMEOPER, 
								GEST_MENU_TB02_FUNZIONI_TMP.ID, "->" 	  	AS [CampiTxt],
								GEST_MENU_TB02_FUNZIONI_TMP.CodFunzioni_s 	AS [CODICE_FUNZIONI],
								GEST_MENU_TB02_FUNZIONI_TMP.Funzioni_s 		AS [FUNZIONI],
								"chiavi" 									AS [TxtKey],
								GEST_MENU_TB02_FUNZIONI_TMP.CodMenu_s 			AS [KeyCodMenu_s]

							FROM 
								GEST_MENU_TB02_FUNZIONI_TMP
								ORDER BY 
								GEST_MENU_TB02_FUNZIONI_TMP.ID
							;

				GEST_MENU_TB02Qry01_01_FUNZIONI_TMP_Delete
					Note
						svuoto la tabelle delle FUNZIONI TMP
					schema
						DELETE 
								GEST_MENU_TB02_FUNZIONI_TMP.*, *
							FROM 
								GEST_MENU_TB02_FUNZIONI_TMP
							WITH 
								OWNERACCESS OPTION;

				GEST_MENU_TB02Qry01_02_INSERT_FUNZIONI_To_TMP
					Note
						accoda le funzioni legate ad un codice Menu
						dalla tabella archivio alla tabella FUNZIONI _TMP
					SCHEMA

						INSERT INTO 
							GEST_MENU_TB02_FUNZIONI_TMP 
							( CodMenu_s, CodFunzioni_s, Funzioni_s, ClassFunzioni_s, NoteFunzioni_s )
							SELECT 
							GEST_MENU_Tb01_VOCI_TMP.CodMenu_s, 
							GEST_MENU_TB02_FUNZIONI.CodFunzioni_s, 
							GEST_MENU_TB02_FUNZIONI.Funzioni_s, 
							GEST_MENU_TB02_FUNZIONI.ClassFunzioni_s, 
							GEST_MENU_TB02_FUNZIONI.NoteFunzioni_s
							FROM 
							(GEST_MENU_Tb01_VOCI_TMP 
							LEFT JOIN 
							GEST_MENU_TB02_FUNZIONI 
							ON GEST_MENU_Tb01_VOCI_TMP.CodMenu_s = GEST_MENU_TB02_FUNZIONI.CodMenu_s) 
							LEFT JOIN 
							GEST_MENU_TB02_FUNZIONI_TMP 
							ON GEST_MENU_TB02_FUNZIONI.CodFunzioni_s = GEST_MENU_TB02_FUNZIONI_TMP.CodFunzioni_s
							WHERE (((GEST_MENU_TB02_FUNZIONI_TMP.CodFunzioni_s) Is Null))
							WITH OWNERACCESS OPTION;




					
			
				
	FUNZIONAMENTO_DEI_MENU	
		Note
			per i menu occorre applicare questo approccio:

			a) per ogni menu selezionato popolare la tabella
				GEST_MENU_Tb01_VOCI_TMP
		   
		    b) per ogni menu selezionato viene popolata la tabella delle
		    	funzioni tmp 

		    c) le due tabelle quella delle voci tmp e quella della funzioni tmp
		    	sono collegate tramite i campo 
		    		CodMenu_s


    		d) alla scelta di un menu nella sottoform_01 viene polato la 
    			combinata 1 e 2 della form master


			1) copio le tabelle tmp progetto 
				MSYS_N01_OGGETTI.mdb
					GEST_MENU_Tb01_VOCI
					GEST_MENU_Tb01_VOCI_TMP
					GEST_MENU_TB02_FUNZIONI
					GEST_MENU_TB02_FUNZIONI_TMP

					CON LE QUERY E LE FORM
					GEST_MENU_Tb01Frm01_M01_GE_VOCI
					GEST_MENU_Tb01Frm01_S01_VOCI
					GEST_MENU_Tb01Frm01_S01_VOCI_TMP	


			2) INDICE_CASELLA_COMBINATA
				attivo la scelta del menu ed aggiorno la combinata 02 della master







                '>>CodMenu   Menu
                Select Case par_Funzione_s
                    'GeCasaMENU  GESTIONE CASA
                    Case ""
                    'GeCondomMENU    Gestione Condomini
                    Case ""
                    'LLPPGestMENU GESTIONE_LAVORI_PUBBLICI
                    Case ""
                    'LTT_MENU    LTT ESTRAZIONI
                    Case ""
                    'MSysMENU    PROGETTO MSys
                    Case ""
                    'Pres3000_Menu GESTIONE_PRESENZE
                    'RUST_Menu   GESTIONE RUST
                    Case ""
                    Case Else
                        MsgBox "NESSUNA ATTIVITA E' ASSOCIATA PER QUESTA FUNZIONE: " & par_Funzione_s
                    End Select
                
                    
                    
        '//............................................................................//
