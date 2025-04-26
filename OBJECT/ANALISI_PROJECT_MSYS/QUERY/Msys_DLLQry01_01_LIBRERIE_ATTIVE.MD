Msys_DLL_LIBRERIEQry01_01_LIBRERIE_ATTIVE.MD


	Note
		query di estrazione delle librerie attive salvate e che riguardano il
		progetto corrente.
		codice ----> @QUERY@Msys_DL_(LE DLL ATTIVE NEL PROGETTO CORRENTE)


	schema
			SELECT 
				Msys_DLL_LIBRERIE.*, 
				"Query delle librerie DLL ATTIVE DEL PROGETTO CORRENTE - Msys_DLL_LIBRERIEQry01_01_LIBRERIE_ATTIVE" AS T
				FROM 
				Msys_DLL_LIBRERIE;


		




