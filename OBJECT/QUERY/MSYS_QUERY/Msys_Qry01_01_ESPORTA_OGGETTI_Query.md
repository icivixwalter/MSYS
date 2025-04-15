NOTA_QUERY:
	Query di esportazione DEGLI OGGETTI QUERY MSYS

	@query@ESPORTA@OGGETTI@QUERY_(la query estrae GLI OGGETTI QUERY @MSYS da @esportare)



SELECT
	
	Msys_Qry01_01_ESPORTA_OGGETTI_Query.md
		Note
			visualizzo le query da esportare estratte dall'archivio


	Msys_Qry01_01_ESPORTA_OGGETTI_Query_TMP.md

		Note
			visualizzo le query da esportare estratte dall'archivio con CODICE PROGETTO
			DA TMP

		SCHEMA_SELECT_IN_TMP
		
				SELECT 
					Msys_ESPORTA_DB_EST.PATH_s, 
					Msys_ESPORTA_ELENCO_OGGETTI_DB.NRO_OGGETTO_i, 
					Msys_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s, 
					Msys_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s, 
					Msys_ESPORTA_DB_EST.Scel_b

					FROM 
						Msys_ESPORTA_DB_EST_TMP 
						INNER JOIN 
							(Msys_ESPORTA_DB_EST INNER JOIN Msys_ESPORTA_ELENCO_OGGETTI_DB 
							ON 
							(Msys_ESPORTA_DB_EST.COD_PROGETTO_s = Msys_ESPORTA_ELENCO_OGGETTI_DB.COD_PROGETTO_s) 
							AND 
							(Msys_ESPORTA_DB_EST.PATH_s = Msys_ESPORTA_ELENCO_OGGETTI_DB.PATH_s)) 
							ON 
							(Msys_ESPORTA_DB_EST.PATH_s = Msys_ESPORTA_DB_EST_TMP.PATH_s) 
							AND 
							(Msys_ESPORTA_DB_EST_TMP.COD_PROGETTO_s = Msys_ESPORTA_DB_EST.COD_PROGETTO_s)

					WHERE 
							(((Msys_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s)="QUERY") 
						AND 
							((Msys_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s)>""))

					WITH 
						OWNERACCESS OPTION;