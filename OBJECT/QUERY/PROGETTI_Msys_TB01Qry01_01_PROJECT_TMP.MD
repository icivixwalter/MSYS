PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP.MD
	Nota
		query creata per l'estrazione dei dati tmp vedi:
			2024_08_23_PROJECT_TMP



	SCHEMA
		
		SELECT 
				PROGETTI_Msys_TB01_PROJECT_TMP.*, 
				"--->" AS TxtChiavi, 
				PROGETTI_Msys_TB01_PROJECT_TMP.COD_PROGETTO_s AS KeyCOD_PROGETTO_s

			FROM 
				PROGETTI_Msys_TB01_PROJECT_TMP;
