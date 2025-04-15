SottoFormMaster_Qry01_01_GE_Msys.Md
	
	Note
		@QUERY@SOTTOFOR@MASTER_(QUERY di estrazione dei progetti relativi alla SOTTOFORM MASTER)

		Query di visualizzazione dei progetti per le sottoform

	SCHEMA


		SELECT 
				"SottoFormMaster_Qry01_01_GE_Msys" AS TxtQuery, 
				SottoFormMaster_Tb01_GE_Msys.COD_PROGETTO_s AS KeyCOD_PROGETTO_s, 
				"->" AS TxtChiavi, 
				SottoFormMaster_Tb01_GE_Msys.*
			FROM 
				SottoFormMaster_Tb01_GE_Msys;
