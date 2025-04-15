MSys_DF01_COMANDI_Qry01_01_SELECT_TUTTI.md
	
	Note
		query di estrazione di tutti i comandi assegnati ad ogni progetto


	schema
		SELECT 
			MSys_DF01_COMANDI.COD_PROGETTO_s, 
			MSys_DF01_COMANDI.PROGETTO_s, 
			MSys_DF01_COMANDI.COMANDO_s, 
			MSys_DF01_COMANDI.Note_cmd_s, 
			MSys_DF01_COMANDI.TIPOGGETTO_s, 
			MSys_DF01_COMANDI.CodCmd_s, 
			MSys_DF01_COMANDI.ID_lng
			
			FROM 
				MSys_DF01_COMANDI;
