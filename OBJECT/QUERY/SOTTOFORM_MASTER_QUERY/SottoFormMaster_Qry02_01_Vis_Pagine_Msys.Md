SottoFormMaster_Qry02_01_Vis_Pagine_Msys.Md
	
	NOTE
		@QUERY@SOTTOFOR@ELEMENTI@MASTER_(QUERY di estrazione DEGLI ELEMENTI relativi alla SOTTOFORM MASTER)

		La query estrare gli elementi del progetto delle sottoform master ed è collegata alla query 


	SCHEMA

		SELECT 
			"SottoFormMaster_Qry02_01_Vis_Pagine_Msys" AS TxtQuery, 
			SottoFormMaster_Tb02_Vis_Pagine_Msys.*, 
			"-->" AS TxtChiavi, 
			SottoFormMaster_Tb02_Vis_Pagine_Msys.COD_PROGETTO_s AS KeyCOD_PROGETTO_s, 
			SottoFormMaster_Tb02_Vis_Pagine_Msys.DENOM_PROGETTO_s AS KeyDENOM_PROGETTO_s, 
			"-->" AS TxtORDi, 
			SottoFormMaster_Tb02_Vis_Pagine_Msys.NRO_OGGETTO_i AS ORD_NRO_OGGETTO_i

			FROM 
				SottoFormMaster_Tb02_Vis_Pagine_Msys
			ORDER BY 
				SottoFormMaster_Tb02_Vis_Pagine_Msys.NRO_OGGETTO_i;
