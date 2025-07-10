PRES3000_QryTb02_01_CalendarioPianoFerie.md

	Note
		query di gestione del calendario piano ferie
			CODICE ---> @PRES3000_QryTb02_01

		Vengono utilizzate le seguenti query di accodamento per gestire il piano ferie dell'ufficio es:


	Name
		PRES3000_QryTb02_}---------------------------------------------@
		PRES3000_QryTb02_01_CalendarioPianoFerie



	SCHEMA

			SELECT 
				PRES3000_Tb02_CalendarioPianoFerie.BLOCC_b, 
				PRES3000_Tb02_CalendarioPianoFerie.HELP_COD_s, 
				PRES3000_Tb02_CalendarioPianoFerie.DENOMINAZIONE_s, 
				PRES3000_Tb02_CalendarioPianoFerie.DataCalend_d, 
				PRES3000_Tb02_CalendarioPianoFerie.GiornoCalend_s, 
				PRES3000_Tb02_CalendarioPianoFerie.AnnoCalend_i, 
				PRES3000_Tb02_CalendarioPianoFerie.MeseCalend_i, 
				PRES3000_Tb02_CalendarioPianoFerie.DESCRIZ_GIORN_s, 
				PRES3000_Tb02_CalendarioPianoFerie.Cod_BrevGiorn_s, 
				PRES3000_Tb02_CalendarioPianoFerie.NoteGiorn_s, 
				PRES3000_Tb02_CalendarioPianoFerie.DATAINS, 
				PRES3000_Tb02_CalendarioPianoFerie.DATAAGG, 
				PRES3000_Tb02_CalendarioPianoFerie.TIMEOPER, 
				PRES3000_Tb02_CalendarioPianoFerie.ID, 
				PRES3000_Tb02_CalendarioPianoFerie.ID AS ID_KEY, "-->campi Form " AS TxtFrm, 
				PRES3000_Tb02_CalendarioPianoFerie.DataCalend_d AS DATA_CALEND, 
				PRES3000_Tb02_CalendarioPianoFerie.GiornoCalend_s AS GIORNO_CALEND, 
				PRES3000_Tb02_CalendarioPianoFerie.MeseCalend_i AS MESE_CALEND, 
				PRES3000_Tb02_CalendarioPianoFerie.Cod_BrevGiorn_s AS COD_BREV_GIORN, 
				PRES3000_Tb02_CalendarioPianoFerie.NoteGiorn_s AS NOTE_GIORN, "--->" AS TxtKey, 
				PRES3000_Tb02_CalendarioPianoFerie.AnnoCalend_i AS Anno_Key, 
				PRES3000_Tb02_CalendarioPianoFerie.MeseCalend_i AS MESE_NUM_key
				FROM PRES3000_Tb02_CalendarioPianoFerie
				ORDER BY 
				PRES3000_Tb02_CalendarioPianoFerie.DENOMINAZIONE_s, 
				PRES3000_Tb02_CalendarioPianoFerie.DataCalend_d, 
				PRES3000_Tb02_CalendarioPianoFerie.DataCalend_d, 
				PRES3000_Tb02_CalendarioPianoFerie.AnnoCalend_i;
