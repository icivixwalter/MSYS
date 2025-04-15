MSys_TUTTI_PROG_VIS_PAGINE_SELECT_Qry01_01_Carica.md

	note
		query di visualizzazione DELLE FORM dei progetti



	schema

		SELECT 
				MSys_TUTTI_PROG_VIS_PAGINE.*, 
				"-->" AS TxtChiavi,

				[MSys_TUTTI_PROG_VIS_PAGINE].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s,

				"-->" AS TxtORDi, 
				[MSys_TUTTI_PROG_VIS_PAGINE].[NRO_OGGETTO_i] AS ORD_NRO_OGGETTO_i
			FROM 
				MSys_TUTTI_PROG_VIS_PAGINE
			ORDER BY 
				[MSys_TUTTI_PROG_VIS_PAGINE].[NRO_OGGETTO_i];


