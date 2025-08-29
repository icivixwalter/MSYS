# TUTTO DA SISTEMARE E RISTUDIARE
	IL PROGETTO DEVE ESSERE CONTROLLATO NEI LINK  e nel suo funzionamento.



## MSYS_OBJECT_DB_EST_ANALISI.MD

	Note
		MSYSDbEst_GE_OBJECT = PROGETTO DATABASE PER LA GESTIONE DEGLI 
		OGGETTI ESTERNI A QUELLO CORRENTE.
		Il progetto di gestione degli oggetti esterni al db corrente che vengono
		gestiti con in COMPONENTI COM, utilizza  l'oggetto application.
		con in questo tutorial:

			start "apri tutorial" call "C:\CASA\LINGUAGGI\ACCESS\TUTORIAL\ACCESS_TUTORIAL\ACCESS_OGETTO_CreateObject_(Esempio_CreateObject_GetObject).pdf"
		
		
		ATTIVITA
			il progetto ha la finalita di aprire db esterni access, visualizzare
			tutti gli oggetti e salvali in una tabella per poi cancellarli o
			ricrearli oppure manipolarli. Gli oggetti base del progetto possono
			essere caricati un nuovo progetto Mdb per cancellare gli oggetti 
			correnti e poi importarli ex novo da progetti base esterni, in
			modo da avere un un progetto generale tutti gli oggetti funzionanti
			creati e testati in piccoli database, in quanto access non permette
			l'utilizzo di multiprogetto.


			CODICE
			Per identificare gli object del progetto di gestione del database
			esterno viene creato il codice univo per tutti gli oggetti, query
			tabelle, form ecc:

				MsysDbEst 			= @codice_(base che identifica il progetto
									 GESTIONE DATABASE ESTERNI)

				Tutti i successivi oggetti devono essere qualificati con il
				codice di appartenenza ad una tabella oppure che non sono legati
				a nessua tabella base. 

				OGGETTI CHE APPARTENGONO ALLA TABELLA 01 = che vengono qualificati dal codice base della tabella di origine che in 
				questo caso è MsysDbEstTb01, per identifare la tabella a cui
				sono legati gli oggetti ossia la tabella 1. Se gli oggetti non
				sono legati a nessuna tabella ma gestiscono procedure collettive
				valide per tutte le tabelle ed oggetti vengono qualificati con
				il codice Tb00 = nessuna tabella di legame.

				MsysDbEstTb01_		= Tabella  (codice di base per gli altri 
									  oggetti)

				query, form report classi e moduli definiti con il codice
				MsysDbEstTb01 = perchè appartengono alla tabella 1 e cosi via.

				@query@codice
				MsysDbEstTb01Qry01_Slt01	= query SELECT
				MsysDbEstTb01Qry01_Dlt01	= query delete
				MsysDbEstTb01Qry01_Ins01	= query Insert
				MsysDbEstTb01Qry01_Upd01	= query UPDATE
				MsysDbEstTb01Qry01_Unn01	= query UNION

				@form@master@codice
				MsysDbEstTb01Frm01_Mts01	= form Master
				MsysDbEstTb01Frm01_Stf01	= form SOTTOFORM 01
				MsysDbEstTb01Rpt01_Mts01	= report Master
				MsysDbEstTb01Rpt01_Stf01	= report sottoreport 01
				
				@moduli@codice
				MsysDbEstTb01Mdl01_Num01	= moduli Di gestione per la tabella 1
				
				@classi@codice
				MsysDbEstTb01Cls01_Num01	= classi legate alla tabella 1
				
				OGGETTI CHE NON APPARTENGONO A NESSUNA TABELLA = allora le query, form. report classi e moduli definiti con il codice
				MsysDbEstTb00 = NON APPARTENGONO A NESSUNA TABELLA
				e cosi via...

				MsysDbEstTb00Mdl01_Num01	= moduli Di gestione per la tabella 1
				
				MsysDbEstTb00Cls01_Num01	= classi legate alla tabella 1
				

				MsysDbEstTb00Qry01_Slt01		= query SELECT

				MsysDbEstTb00Frm01_Stf01		= form SOTTOFORM 01

				MsysDbEstTb00Mdl01_Num01		= moduli Di gestione valido per 
											     tutte le tabelle
				
				MsysDbEstTb00Cls01_Num01	   = classi non legate alla tabella 1 ma 
										        valide per tutte le tabelle

				MsysDbEstTb00Mcr01_N01	   		= Le macro



		ATTIVITA:  la funzione del progetto è quella di accedere ai db esterni e
		visualizzare controllorare gli oggetti del db esterno.



		A) VISUALIZAZIONE = con la costruzione delle query e delle tabelle e le form imposto le basi per la visualizzazione degli oggetti esterni.

	TABELLE
		Note
			le tabelle create sono 

		MsysDbEstTb00_}------------------------------------------------@
		MsysDbEstTb01_OBJECT
			Note
				la chiave primaria è il campo 
					Name1_s 			= keyprimary campo nome tabella
					

		MSysTb01_OBJECT_DB_EST_TMP
			Nota
				tabella e quella tmp che contiene tutti gli oggetti salvati ed esistenti nel db esterno esaminato
				La chiavie primaria è:

					Name1_s 			= keyprimary campo nome tabella

		MSysTb05_DB_EST
			Note
				tabella per identificar il db esterno da monitorare

	QUERY
		SELECT
			MsysDbEstTb01Qry01_Slt00_}-------------------------------------@
			MsysDbEstTb01Qry01_Slt01_OBJECT

					SELECT	
						MsysDbEstTb01_OBJECT_TMP.NRO_OGGETTO_i, 
						MsysDbEstTb01_OBJECT_TMP.TIPOGGETTO_s, 
						MsysDbEstTb01_OBJECT_TMP.NOMEOGGETTO_s, 
						MsysDbEstTb01_OBJECT_TMP.NOTA_OGGETTO_s, 
						MsysDbEstTb01_OBJECT_TMP.NOTEex_s, 
						MsysDbEstTb01_OBJECT_TMP.EXPORT_b, 
						MsysDbEstTb01_OBJECT_TMP.COD_PROGETTO_s, 
						MsysDbEstTb01_OBJECT_TMP.DISCO_s, 
						MsysDbEstTb01_OBJECT_TMP.PATH_s, 
						MsysDbEstTb01_OBJECT_TMP.DATABASE_s, 
						MsysDbEstTb01_OBJECT_TMP.Attribute_Lng, 
						MsysDbEstTb01_OBJECT_TMP.Expression_m, 
						MsysDbEstTb01_OBJECT_TMP.Flag_b, 
						MsysDbEstTb01_OBJECT_TMP.LvExtra_i, 
						MsysDbEstTb01_OBJECT_TMP.Name1_s, 
						MsysDbEstTb01_OBJECT_TMP.Name2_s, 
						MsysDbEstTb01_OBJECT_TMP.ObjectId_lng, 
						MsysDbEstTb01_OBJECT_TMP.Order_b, 
						MsysDbEstTb01_OBJECT_TMP.NOTE_ATTRIBUTI_s, 
						MsysDbEstTb01_OBJECT_TMP.PROPRIETA_s, 
						MsysDbEstTb01_OBJECT_TMP.VALOREPROPRIETA_m, 
						MsysDbEstTb01_OBJECT_TMP.COMANDO_s, 
						MsysDbEstTb01_OBJECT_TMP.Scel_b, 
						MsysDbEstTb01_OBJECT_TMP.Stam_b, 
						MsysDbEstTb01_OBJECT_TMP.Canc_b, 
						MsysDbEstTb01_OBJECT_TMP.Agg_b, 
						MsysDbEstTb01_OBJECT_TMP.BLOCC_b, 
						MsysDbEstTb01_OBJECT_TMP.IDKeyArch_lng, 
						MsysDbEstTb01_OBJECT_TMP.Msg01_s, 
						MsysDbEstTb01_OBJECT_TMP.DATAINS_d, 
						MsysDbEstTb01_OBJECT_TMP.DATAAGG_d, 
						MsysDbEstTb01_OBJECT_TMP.TIMEOPER_d, 
						MsysDbEstTb01_OBJECT_TMP.ORA_AGG_d, 
						MsysDbEstTb01_OBJECT_TMP.ID_TAB_lng
						FROM 
							MsysDbEstTb01_OBJECT_TMP
						ORDER BY 
							MsysDbEstTb01_OBJECT_TMP.ID_TAB_lng
						WITH 
							OWNERACCESS OPTION;

				Note
					query suglio oggetti del db esterno caricati

				
			MsysDbEstTb01Qry01_Slt01_OBJECT_TMP
			

			
			MsysDbEstTb05Qry01_Slt00_}-------------------------------------@
			MsysDbEstTb05Qry01_Slt01_DB_ESTERNO
				Note
					@query@db@esterno_(Visualizzo le path del db esterno da esaminare)

				codice
						@MsysDbEstTb05Qry01_(@qry@db o query per il db esterno)
				schema
					SELECT 
						MSysTb05_DB_EST.DISCO_s, 
						MSysTb05_DB_EST.PATH_s, 
						MSysTb05_DB_EST.DATABASE_s, 
						MSysTb05_DB_EST.Scel_b, 
						MSysTb05_DB_EST.Stam_b, 
						MSysTb05_DB_EST.Canc_b, 
						MSysTb05_DB_EST.Agg_b, 
						MSysTb05_DB_EST.BLOCC_b, 
						MSysTb05_DB_EST.DATAINS_d, 
						MSysTb05_DB_EST.DATAAGG_d, 
						MSysTb05_DB_EST.TIMEOPER_d, 
						MSysTb05_DB_EST.ORA_AGG_d, 
						MSysTb05_DB_EST.ID_DB_EST_lng 
						FROM 
							MSysTb05_DB_EST
						WITH 
							OWNERACCESS OPTION;


		DELETE
			MsysDbEstTb01Qry01_Dlt00_}-------------------------------------@
			MsysDbEstTb01Qry01_Dlt01_OBJECT
				Note
					svuoto la tabella di archivio deglio oggetti importati

				SCHEMA

					DELETE 
							MsysDbEstTb01_OBJECT.*, *
						FROM 
							MsysDbEstTb01_OBJECT
						WITH 
							OWNERACCESS OPTION;


			MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP
				Note
					svuoto la tabella TMP deglio oggetti

				SCHEMA

					DELETE 
							MsysDbEstTb01_OBJECT_TMP.*, *
						FROM 
							MsysDbEstTb01_OBJECT_TMP
						WITH 
							OWNERACCESS OPTION;


		INSERT
			MsysDbEstTb01Qry01_Ins00_}-------------------------------------@
			MsysDbEstTb01Qry01_Ins01_OBJECT_Tmp_To_Archivio
				Note
					Accodo tutti gli oggetti tabella recuperati nel db esterno nella tabella
					di archivio 
					@accodo@tabelle_(@da@tmp in archivio tabelle database esterno)

				schema
					
					INSERT 
						INTO MsysDbEstTb01_OBJECT 
						( NRO_OGGETTO_i, TIPOGGETTO_s, NOMEOGGETTO_s, NOTA_OGGETTO_s, NOTEex_s, EXPORT_b, COD_PROGETTO_s, DISCO_s, PATH_s, DATABASE_s, Attribute_Lng, Expression_m, Flag_b, LvExtra_i, Name1_s, Name2_s, ObjectId_lng, Order_b, NOTE_ATTRIBUTI_s, PROPRIETA_s, VALOREPROPRIETA_m, COMANDO_s, Scel_b, Stam_b, Canc_b, Agg_b, BLOCC_b, IdOrd_lng, IDKeyArch_lng, Msg01_s, DATAINS_d, DATAAGG_d, TIMEOPER_d, ORA_AGG_d )

							SELECT 
								MsysDbEstTb01_OBJECT_TMP.NRO_OGGETTO_i, 
								MsysDbEstTb01_OBJECT_TMP.TIPOGGETTO_s, 
								MsysDbEstTb01_OBJECT_TMP.NOMEOGGETTO_s, 
								MsysDbEstTb01_OBJECT_TMP.NOTA_OGGETTO_s, 
								MsysDbEstTb01_OBJECT_TMP.NOTEex_s, 
								MsysDbEstTb01_OBJECT_TMP.EXPORT_b, 
								MsysDbEstTb01_OBJECT_TMP.COD_PROGETTO_s, 
								MsysDbEstTb01_OBJECT_TMP.DISCO_s, 
								MsysDbEstTb01_OBJECT_TMP.PATH_s, 
								MsysDbEstTb01_OBJECT_TMP.DATABASE_s, 
								MsysDbEstTb01_OBJECT_TMP.Attribute_Lng, 
								MsysDbEstTb01_OBJECT_TMP.Expression_m, 
								MsysDbEstTb01_OBJECT_TMP.Flag_b, 
								MsysDbEstTb01_OBJECT_TMP.LvExtra_i, 
								MsysDbEstTb01_OBJECT_TMP.Name1_s, 
								MsysDbEstTb01_OBJECT_TMP.Name2_s, 
								MsysDbEstTb01_OBJECT_TMP.ObjectId_lng, 
								MsysDbEstTb01_OBJECT_TMP.Order_b, 
								MsysDbEstTb01_OBJECT_TMP.NOTE_ATTRIBUTI_s, 
								MsysDbEstTb01_OBJECT_TMP.PROPRIETA_s, 
								MsysDbEstTb01_OBJECT_TMP.VALOREPROPRIETA_m, 
								MsysDbEstTb01_OBJECT_TMP.COMANDO_s, 
								MsysDbEstTb01_OBJECT_TMP.Scel_b, 
								MsysDbEstTb01_OBJECT_TMP.Stam_b, 
								MsysDbEstTb01_OBJECT_TMP.Canc_b, 
								MsysDbEstTb01_OBJECT_TMP.Agg_b, 
								MsysDbEstTb01_OBJECT_TMP.BLOCC_b, 
								MsysDbEstTb01_OBJECT_TMP.IdOrd_lng, 
								MsysDbEstTb01_OBJECT_TMP.IDKeyArch_lng, 
								MsysDbEstTb01_OBJECT_TMP.Msg01_s, 
								MsysDbEstTb01_OBJECT_TMP.DATAINS_d, 
								MsysDbEstTb01_OBJECT_TMP.DATAAGG_d, 
								MsysDbEstTb01_OBJECT_TMP.TIMEOPER_d, 
								MsysDbEstTb01_OBJECT_TMP.ORA_AGG_d
								
							FROM 
								MsysDbEstTb01_OBJECT_TMP LEFT JOIN 
								MsysDbEstTb01_OBJECT ON 
								MsysDbEstTb01_OBJECT_TMP.Name1_s = MsysDbEstTb01_OBJECT.Name1_s
								

							WHERE 
								(((MsysDbEstTb01_OBJECT.Name1_s) Is Null))
							WITH 
								OWNERACCESS OPTION;

	FORM
		MSysTb01Frm00_{@==============================================@}
		
		MsysDbEstTb01Frm01_Mts01_}-------------------------------------@
		MsysDbEstTb01Frm01_Mts01_GE_OBJECT	= form Master
			Note
				form master per la gestione degli oggetti esterni
				questa form carica gli oggetti esterni solo se ci troviamo
				in un db diverso da quello di origine
		
		MsysDbEstTb01Frm01_Stf01_OBJECT
			Note
				LE FORM degli oggetti del db esterno caricati in tabella.
				Questi oggetti potranno essere utilizzati per la maninopolazione.
				Il funzionamento, apri il db esterno controlli le form 
				e carichi i nome su una tabella tmp
		
		MsysDbEstTb01Frm01_Stf01_OBJECT_TMP
			Note
				la form che visualizza solo gli oggetti controllati nel db esterno
				e salvati in tmp
		
		MsysDbEstTb05Frm01_Mt00_}--------------------------------------@
		MsysDbEstTb05Frm01_Stf01_DB_ESTERNO
			Note
				oggetti visualizzati nel db esterno

			Proprieta
				recordSource = @MsysDbEstTb05Qry01_(codice @query inserita nella form 
													per il @db@esterno)

	


		UTILITA_Frm12_SERVIZI_TUTTI
			Note
				form che contiene i comandi dei servizi
				01) cancella
				02) visulizza dati db esterno

		UTILITA_MsysDbEst_Frm20_}--------------------------------------@
		UTILITA_MsysDbEst_Frm20_IMPORTA_PROGETTO
			Note
				todo: ho creato un'altra form di importazione del progetto separata

				questa form deve prevedere il caricamento degli oggetti in un'altro
				progetto form e quindi carica tutti gli oggetti del progetto My sys
				database esterno. 
				Attenzione se si trova nel progetto originale cancella le form originale
				impostare un controllo del progetto per evitare conflitti.
				Il progetto originale si chiama Project_MSYSDbEst_GE_OBJECT e quindi creare una funzione che impedisce il caricamento se ci troviamo in questo progetto.

				CONTROLLO DEL PROGETTO: per controllare se ci troviamo nel progetto base è stata
							creata la funzione CheckProjectName_b_PFunct la quale controllo se ci troviamo
							nel progetto MSYSDbEst_GE_OBJECT.mdb ed impedisce il caricamento degli oggetti
							table, query ecc. per evitare sovrapposizioni.

	MODULI
		MsysDbEstTb01Mdl00_}-------------------------------------------@
		MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO
			nota:
				il modulo attiva il controllo presso il database esterno
				e per il momento le tabelle.
					@modulo_(per il @controllo@oggetti del @db@esterno)



			LINK
				per aprire la IL MODULO per la gestione dei databse esterni 
				attivare questo link:
					start "apri modulo gestione db esterno" call "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MDB\OBJECT\FORM\MsysDbEstTb01Mdl01_Num01_CONTROLLO_DB_ESTERNO.bas"

	MACRO
		Note
			ho creato le macro per il salvataggio su xls dei parametri di gestione
			dell'oggetto.

		MsysDbEstTb00Mcr01_N01_ESPORTA_XLS
			Note
				macro che esporta i parametri di gestione su excel


		MsysDbEstTb00Mcr01_N02_COLLEGA_XLS
			Note 
				macro che effettua il collegamento da excel dei parametri esportati



# GIORNALE_DELLE_MODIFICHE
	Note
		Le modifiche effettuate sono le seguenti:

	2024
		GIORNALE_OPERAZIONI
		2024
			2024.07.27= Nella TreeView creato 3 Macro di Importazione tabelle

				TREEVIEW_COLLEGA_Mcr00_{@===============================@}
				TREEVIEW_COLLEGA_Mcr01_Msys_DF13_DLL_PROGETTO				= macro che importa la 
																			tabella DF13 dei progetti Dll
				TREEVIEW_COLLEGA_Mcr02_Msys_DF14_DLL_LIBRERIE				= macro che importa la tabella dll DF14 dove sono salvate
																			le librerie collegate ai progetti
				TREEVIEW_COLLEGA_Mcr03_Msys_DLL								= macro che importa la tabella in cui vengono salvate le
																			librerie utilizzate nel progetto corrente e che serve per 
																			aggiornare o accodare le librerie nella DF13 E DF14

			2024.07.27= Nella TreeView ho inserito

				UTILITA_Msys_Mdl00­­_}-----------------------------------------@
				UTILITA_Msys_Mdl03_ACCESS_REFERENZIA_LE_LIBRERIE			= inserita la routine delle librerie che è stata modificata
																			in quanto la stessa scrive le librerie Dll del progetto
																			corrente nella tabella in Msys_DLL
