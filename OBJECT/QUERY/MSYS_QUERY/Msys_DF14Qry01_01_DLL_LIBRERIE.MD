Msys_DF14Qry01_01_DLL_LIBRERIE.MD

	Note
		la query di estrazione delle librerie DLL legate ai progetti.
		codice ----> @QUERY@DF14@DLL@LIBRERIE_(le librerie DLL legate ai progetti)


	schema
			SELECT 
				Msys_DF14_DLL_LIBRERIE.*, 
				"Query delle librerie DLL collegate ai progetti - Msys_DF14Qry01_01_DLL_LIBRERIE" AS T
				FROM 
				Msys_DF14_DLL_LIBRERIE;

				
		




