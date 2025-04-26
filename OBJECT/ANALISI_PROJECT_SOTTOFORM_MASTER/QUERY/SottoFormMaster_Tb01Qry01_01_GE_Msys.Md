SottoFormMaster_Tb01Qry01_01_GE_Msys_TMP.Md
	
	Note
			CODICE ----> @SottoFormMaster_Qry01_01_GE_Msys@QUERY@SOTTOFOR@MASTER_(QUERY di estrazione dei progetti relativi alla SOTTOFORM MASTER)

		Query di visualizzazione dei progetti per le sottoform CON IL 

	SCHEMA

			SELECT 
					"@SottoFormMaster_Qry01_01_GE_Msys" AS TxtQuery, 
					SottoFormMaster_Tb01_GE_Msys_tmp.NRO_PROGETTO_i, 
					SottoFormMaster_Tb01_GE_Msys_tmp.COD_PROGETTO_s, 
					SottoFormMaster_Tb01_GE_Msys_tmp.DENOM_PROGETTO_s, 
					SottoFormMaster_Tb01_GE_Msys_tmp.NRO_OGGETTO_i, 
					SottoFormMaster_Tb01_GE_Msys_tmp.COD_FORMS_s, 
					SottoFormMaster_Tb01_GE_Msys_tmp.NOTE_PROGETTO_s, 
					SottoFormMaster_Tb01_GE_Msys_tmp.VISIBLE_b, 
					SottoFormMaster_Tb01_GE_Msys_tmp.DATAINS_d, 
					SottoFormMaster_Tb01_GE_Msys_tmp.DATAAGG_d, 
					SottoFormMaster_Tb01_GE_Msys_tmp.TIMEOPER_d, 
					SottoFormMaster_Tb01_GE_Msys_tmp.ORA_AGG_d, 
					SottoFormMaster_Tb01_GE_Msys_tmp.ID, "->" AS TxtChiavi, 
					SottoFormMaster_Tb01_GE_Msys_tmp.COD_PROGETTO_s AS KeyCOD_PROGETTO_s

				FROM 
					SottoFormMaster_Tb01_GE_Msys_tmp;



SottoFormMaster_Tb01Qry01_01_GE_Msys_TMP.md
	
	Note
		query che estrae la tabella sottoform master tmp


	SCHEMA


		SELECT 
			"SottoFormMaster_Tb01Qry01_01_GE_Msys_TMP" AS TxtQuery,
			SottoFormMaster_Tb01_GE_Msys_tmp.COD_PROGETTO_s AS KeyCOD_PROGETTO_s,
			"->" AS TxtChiavi, 
			SottoFormMaster_Tb01_GE_Msys_tmp.*
		FROM 
			SottoFormMaster_Tb01_GE_Msys_tmp;

