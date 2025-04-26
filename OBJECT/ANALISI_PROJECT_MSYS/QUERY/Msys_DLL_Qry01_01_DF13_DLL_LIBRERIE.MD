Msys_DLL_Qry01_01_DF13_DLL_LIBRERIE.MD

	Note
		estraggo i progetti con le libreria DLL necessarie.
			codice ----> @QUERY@DF13@DLL@PROGETTI_(I progetto e le librerie DLL necessarie)



	SCHEMA

		SELECT 
				Msys_DF13_DLL_PROGETTO.*, 
				"Query progetti e DLL - Msys_DF13Qry01_01_DLL_PROGETTO" AS T
			FROM 
				Msys_DF13_DLL_PROGETTO;


		


	PROPRIETA
		Nome foglio secondario		: Query.Msys_DF14Qry01_01_DLL_LIBRERIE
		collega campi master		: COD_DLL_PROGETTO_s
		collega campi secodnari		: COD_DLL_PROGETTO_s