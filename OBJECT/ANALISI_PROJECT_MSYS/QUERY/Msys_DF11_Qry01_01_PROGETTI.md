Msys_DF11_Qry01_}----------------------------------------------@
Msys_DF11_Qry01_01_PROGETTI

	Note
		Query di gestione dei progetti
			@query.gestione.progetti
			@progetti.query


	STRUTTURA

		SELECT 
		Msys_DF11_PROGETTI.COD_PROGETTO_s, 
		Msys_DF11_PROGETTI.PROGETTO_s, 
		Msys_DF11_PROGETTI.Note_PROGETTO_s, 
		"" 									AS [TxtChiavi], 
		Msys_DF11_PROGETTI.COD_PROGETTO_s 	AS [KeyCOD_PROGETTO_s]
		FROM Msys_DF11_PROGETTI;


	PROPRIETA
		fogli dati secondario 	---> QueryMsys_DF12_Qry01_01_TIPO_OGGETTI
		detail					---> KeyCOD_PROGETTO_s
		master					---> KeyCOD_PROGETTO_s	
