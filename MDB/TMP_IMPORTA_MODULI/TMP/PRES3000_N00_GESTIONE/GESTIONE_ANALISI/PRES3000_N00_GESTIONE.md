PRES3000_N00_GESTIONE.md

	PROGETTO
		Note
			@progetto@di@gestione
			Path: c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N00_GESTIONE\
			file: Project_PRES3000_N00_GESTIONE.sublime-workspace



	FORM_MASTER
			Note
				La form @master gestisce tutte le ipotesi della giustificazioni
				contiene un @tab @controll con le seguenti sottoform:
		
		PRES3000_Frm01_M01_SOTTO_FORM_A_SCHEDE
			TAB_CONTROLL
				note
					@tab@control_(per la gestione delle presenze)
					
					Il tab controll per ora non carica in modo automatico e i vari tab contengono le seguenti form

				01
					PRES3000_Frm04_S02_01_RAGGR_ANNO
						Note
							form di raggruppamento anno

			
			
				02
					PRES3000_Frm01_S01_01_Calendario
						Note
						form calendario che si basa sulla tabella omonima PRES3000_Tb01_Calendario recuperata tramite la query 		PRES3000_QryTb0101_01_Calendario
						c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N01_CALENDARIO.mdb;TABLE=PRES3000_Tb01_Calendario


				03
					PRES3000_Frm01_S04_03_RAGGRUPPA_ANNO_GIUSTIFICAZIONI
						Note
							form che raggruppa le giustificazioni per anno tutte
				04
					PRES3000_Frm01_S11_GROUP_GANNO_Giustificazioni_ANNO_CORRENTE
						Note
							form che raggruppa le giustificazioni per anno solo l'anno corrente
				05
				06
				07
				08
				09
					PRES3000_Frm02_04_ElencoGiornaliere_TUTTE_ANNO_CORRENTE
						@dove@si@trova?
				10
					PRES3000_FrmTb25_11_GROUP_Dipendenti_ALF
						@dove@si@trova?

				12 PRES3000_Frm12_01_SERVIZI_TAB_12
					Note
						form relativa ai servizi del pres3000


		



	UTILITA
		CERCA SOTTO PROGETTI
			DEFINIZIONE ORARI
				Note
					@apri@orari@definizione_(la definizione degli orari @timbrature ed altro)
					Si trova nel progetto orari e timbrature in cui si trova anche
					la definizione del calendario annuale delle @festivita da aggiornare ogni anno


				PATH: 
					c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB\
				PROJECT: 
					PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB.mdb

				APRI
						@start@apri@orari_(apri progetto mdb orari)
					start "Apri definizione orari" call "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB\PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB.mdb"








                   
                '//accoda solo le tabelle
                    If TABLE_02_s > "" Then
                        On Error Resume Next        '//se non ci sono errori cancella
                        DoCmd.DeleteObject acTable, TABLE_02_s
                        DoCmd.TransferDatabase acLink, "Microsoft Access", Database_PATH_s & Database_s, acTable, TABLE_02_s, TABLE_02_s, False
                    End If
