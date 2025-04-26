MSys_QUERY_Qry01_01_SELECT_TUTTE_TMP.MD

	Note
		ho creato questa query per estrarre solo le query con il codice
		progetto salvata su progetti TMP


	SCHEMA


		SELECT 
				MSys_QUERY.*, 
				MSys_QUERY.COD_PROGETTO_s AS ORD_COD_PROGETTO_s, MSys_QUERY.NRO_OGGETTO_i AS ORD
			FROM 
				MSys_QUERY INNER JOIN 
				PROGETTI_Msys_TB01_PROJECT_TMP ON 
				MSys_QUERY.COD_PROGETTO_s = PROGETTI_Msys_TB01_PROJECT_TMP.COD_PROGETTO_s
			ORDER BY 
				MSys_QUERY.COD_PROGETTO_s, MSys_QUERY.NRO_OGGETTO_i;
