PRES3000_QryDF01_01_CalendarioFestivita
	
	Note
		Query sul calendario delle attivita
			codice PRES3000_QryDF01_01
				

	SCHEMA
		SELECT 
			PRES3000_DF01_CalendarioFestivita.DataCalend_d, PRES3000_DF01_CalendarioFestivita.GiornoCalend_s, PRES3000_DF01_CalendarioFestivita.AnnoCalend_i, PRES3000_DF01_CalendarioFestivita.NoteGiorn_s, PRES3000_DF01_CalendarioFestivita.Cod_BrevGiorn_s, PRES3000_DF01_CalendarioFestivita.DescrizioneGiorn_s, PRES3000_DF01_CalendarioFestivita.IDFEST, "-->" AS Txtkey, PRES3000_DF01_CalendarioFestivita.AnnoCalend_i AS KeyAnnoCalend_i
			FROM PRES3000_DF01_CalendarioFestivita
			ORDER BY PRES3000_DF01_CalendarioFestivita.DataCalend_d;
