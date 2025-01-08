MSys_DF01_COMANDI_Qry01_02_GROUP_TUTTI.md

	Note
		query di raggruppamento dei comandi per progetto





	SCHEMA

		SELECT 
				MSys_DF01_COMANDI.COD_PROGETTO_s, 
				MSys_DF01_COMANDI.PROGETTO_s, 
				Count(MSys_DF01_COMANDI.COD_PROGETTO_s) AS TOT_c
			FROM 
				MSys_DF01_COMANDI
			GROUP 
				BY MSys_DF01_COMANDI.COD_PROGETTO_s, MSys_DF01_COMANDI.PROGETTO_s;


	PROPRIETA
		Note
			collegata alla query sottostante.
		Nome foglio secondario
			MSys_DF01_COMANDI_Qry01_01_SELECT_TUTTI
		collega campi secondari
			COD_PROGETTO_s;PROGETTO_s
		collega campi master
			COD_PROGETTO_s;PROGETTO_s
