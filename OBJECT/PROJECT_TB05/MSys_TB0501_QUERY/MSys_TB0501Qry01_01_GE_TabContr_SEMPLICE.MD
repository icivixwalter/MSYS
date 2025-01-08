NOTE:
	query per la gestione del TAB CONTROL SEMPLICE il codice progetto


QUERY
	MSys_TB0501Qry00_}---------------------------------------------@
	MSys_TB0501Qry01_01_GE_TabContr_SEMPLICE
		SCHEMA
			SELECT 
				MSys_TB0501_GE_TabContr_SEMPLICE.*, "->" AS TxtChiavi, MSys_TB0501_GE_TabContr_SEMPLICE.COD_PROGETTO_s AS KeyCOD_PROGETTO_s
				FROM MSys_TB0501_GE_TabContr_SEMPLICE;



	MSys_TB0501Qry01_01_GE_TabContr_SEMPLICE_TMP
		SCHEMA
		
			SELECT 
				MSys_TB0501_GE_TabContr_SEMPLICE_TMP.*, "->" AS TxtChiavi, MSys_TB0501_GE_TabContr_SEMPLICE_TMP.COD_PROGETTO_s AS KeyCOD_PROGETTO_s
				FROM MSys_TB0501_GE_TabContr_SEMPLICE_TMP;


								






