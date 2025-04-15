SottoFormMaster_Tb02Qry01_01_Vis_Pagine_Msys.Md
	
	NOTE
		@SottoFormMaster_Tb02Qry01_01_Vis_Pagine_Msys@QUERY@SOTTOFOR@ELEMENTI@MASTER_(QUERY di estrazione DEGLI ELEMENTI relativi alla SOTTOFORM MASTER)

		La query estrare gli elementi del progetto delle sottoform master ed è collegata alla query 

			

	SCHEMA

		SELECT 
			"@SottoFormMaster_Tb02Qry01_01_Vis_Pagine_Msys" AS TxtQuery, 
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




SottoFormMaster_Tb02Qry01_01_GE_Msys_TMP.md
	Note
		creato anche la tabella tmp dove vengon salvate le patine da visualzzare.

		"@SottoFormMaster_Tb02Qry01_01_Vis_Pagine_Msys_TMP@QUERY@SOTTOFOR@ELEMENTI@MASTER_(QUERY di estrazione PAGINE DA VISUALIZZARE salvate nella cartella TMP)

	SCHEMA
		
			SELECT 
					"@SottoFormMaster_Tb02Qry01_01_Vis_Pagine_Msys_TMP" AS TxtQuery, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.NRO_PROGETTO_i, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.COD_PROGETTO_s, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.DENOM_PROGETTO_s, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.NRO_OGGETTO_i, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.TIPOGGETTO_s, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.NOMEOGGETTO_s, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.NOMEpagina_s, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.VISIBLE_b, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.NOTA_OGGETTO_s, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.NOTEex_s, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.DATAINS_d, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.DATAAGG_d, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.TIMEOPER_d, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.ORA_AGG_d, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.ID, 
					"-->" AS TxtChiavi, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.COD_PROGETTO_s AS KeyCOD_PROGETTO_s, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.DENOM_PROGETTO_s AS KeyDENOM_PROGETTO_s, 
					"-->" AS TxtORDi, 
					SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.NRO_OGGETTO_i AS ORD_NRO_OGGETTO_i

				FROM 
						SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP
				ORDER BY 
						SottoFormMaster_Tb02_Vis_Pagine_Msys_TMP.NRO_OGGETTO_i;
