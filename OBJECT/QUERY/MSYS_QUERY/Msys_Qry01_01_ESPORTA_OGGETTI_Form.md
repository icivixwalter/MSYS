NOTA_QUERY:
	Query di esportazione DI TUTTI GLI OGGETTI MACRO MSYS
		@query@ESPORTA@OGGETTI@TUTTI_(la query estrae TUTTI GLI OGGETTI @MSYS da @esportare)

		Vengono esportati le query necessarie per il progetto msys.



SELECT

	Msys_Qry01_01_ESPORTA_OGGETTI_Form.md


	Msys_Qry01_01_ESPORTA_OGGETTI_Form_TMP
		
		SELECT_IN_TMP
			Seleziono la tabella degli oggetti form con il codice uguale alla tabella DB EST TMP
				@SELEZIONA@FORM@TMP_(estraggo solo le form tmp con il codice progetto salvato nella
									tabella tmp)
		



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
							(((Msys_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s)="forms") 
						AND 
							((Msys_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s)>""))

					WITH 
						OWNERACCESS OPTION;
