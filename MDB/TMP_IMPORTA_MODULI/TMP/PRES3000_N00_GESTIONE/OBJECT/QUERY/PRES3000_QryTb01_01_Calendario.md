PRES3000_QryTb01_01_Calendario.md

	
	Note
		@query@calendario_(sul @calendario)


	name
		PRES3000_QryTb01_}---------------------------------------------@
		PRES3000_QryTb01_01_Calendario
		
	schema
			SELECT 
				PRES3000_Tb01_Calendario.*, 
				PRES3000_Tb01_Calendario.ID AS ID_KEY, 
				"-->campi Form " AS TxtFrm, 
				PRES3000_Tb01_Calendario.DataCalend_d AS DATA_CALEND, PRES3000_Tb01_Calendario.GiornoCalend_s AS GIORNO_CALEND, PRES3000_Tb01_Calendario.MeseCalend_i AS MESE_CALEND, PRES3000_Tb01_Calendario.FileCalend_s AS FILE_CALEND, PRES3000_Tb01_Calendario.Cod_BrevGiorn_s AS COD_BREV_GIORN, PRES3000_Tb01_Calendario.CONT_GIORN_i AS CONTGIORN, PRES3000_Tb01_Calendario.TOT_GIORN_i AS TOTGIORN, PRES3000_Tb01_Calendario.NoteGiorn_s AS NOTE_GIORN, PRES3000_Tb01_Calendario.FileCalend_s AS FILE, PRES3000_Tb01_Calendario.PathFile_s AS PATH, "--->" AS TxtKey, PRES3000_Tb01_Calendario.AnnoCalend_i AS Anno_Key, PRES3000_Tb01_Calendario.MeseCalend_i AS MESE_NUM_key
				FROM 
				PRES3000_Tb01_Calendario
				ORDER 
				BY PRES3000_Tb01_Calendario.DataCalend_d, PRES3000_Tb01_Calendario.AnnoCalend_i;
