MSys_TUTTI_PROG_INSERT_Qry01_01_PROGETTI_in_TMP.md


	Note
		Query di accodamento del progetto selezionato nella tabella tmp PROGETTI
		Esempio di query per l'inserimento dei codice dei progetti.


	schema
		
		INSERT 
			INTO 
				MSys_TUTTI_PROG_TMP ( NRO_OGGETTO_i, COD_PROGETTO_s, DENOM_PROGETTO_s )
			SELECT 
				MSys_TUTTI_PROG.NRO_OGGETTO_i, 
				MSys_TUTTI_PROG.COD_PROGETTO_s, 
				MSys_TUTTI_PROG.DENOM_PROGETTO_s
				FROM MSys_TUTTI_PROG 
				LEFT JOIN MSys_TUTTI_PROG_TMP 
				ON (MSys_TUTTI_PROG.DENOM_PROGETTO_s = MSys_TUTTI_PROG_TMP.DENOM_PROGETTO_s) 
				AND (MSys_TUTTI_PROG.COD_PROGETTO_s = MSys_TUTTI_PROG_TMP.COD_PROGETTO_s)
				WHERE 
					(((MSys_TUTTI_PROG.COD_PROGETTO_s)="MSys") 
					AND 
					((MSys_TUTTI_PROG_TMP.NRO_OGGETTO_i) Is Null) AND ((MSys_TUTTI_PROG_TMP.COD_PROGETTO_s) Is Null));




					