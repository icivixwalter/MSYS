QUERY ESPORTA OGGETTI
	Ho creato diverse query tra cui quelle piu significative, la query di estrazione degli oggetti
	che è stata ridenominata da Msys_Qry01_01_ESPORTA_OGGETTI ---> Msys_Qry01_01_ESPORTA_OGGETTI_TUTTI
	dalla esporta oggetti tutti è stata ricavata la query
		Msys_Qry01_01_ESPORTA_OGGETTI_TMP
	che serve ad estrarre solo quelli con il codice progetto salvato in tmp.

SELECT

	Msys_Qry01_01_ESPORTA_OGGETTI.md
	Nota
		nuova query che deve gestire la nuova esportazione degli oggetti.



	SCHEMA

		SELECT 
			Msys_ESPORTA_ELENCO_OGGETTI_DB.NRO_OGGETTO_i, 
			Msys_ESPORTA_ELENCO_OGGETTI_DB.COD_PROGETTO_s, 
			Msys_ESPORTA_ELENCO_OGGETTI_DB.PATH_s, 
			Msys_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s, 
			Msys_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s

			FROM 
				Msys_ESPORTA_ELENCO_OGGETTI_DB 
			ORDER 
				BY Msys_ESPORTA_ELENCO_OGGETTI_DB.NRO_OGGETTO_i 
			WITH 
				OWNERACCESS OPTION;





				



	Msys_Qry01_01_ESPORTA_OGGETTI_TMP
		questa query estrae solo gli oggetti che hanno il codice progetto salvato in tmp. In questo modo
		si rende compatibile la visualizazione degli oggetti da esportare con la path del db scelto.

		SCHEMA

			SELECT 
				Msys_ESPORTA_ELENCO_OGGETTI_DB.NRO_OGGETTO_i, 
				Msys_ESPORTA_ELENCO_OGGETTI_DB.COD_PROGETTO_s, 
				Msys_ESPORTA_ELENCO_OGGETTI_DB.PATH_s, 
				Msys_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s, 
				Msys_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s

				FROM 
					Msys_ESPORTA_ELENCO_OGGETTI_DB 
				ORDER 
					BY Msys_ESPORTA_ELENCO_OGGETTI_DB.NRO_OGGETTO_i 
				WITH 
					OWNERACCESS OPTION;