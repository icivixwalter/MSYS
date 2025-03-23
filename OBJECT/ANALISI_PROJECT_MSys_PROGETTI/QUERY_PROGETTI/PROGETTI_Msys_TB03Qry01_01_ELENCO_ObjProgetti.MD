PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti.MD

Note
		@query@ELENCO@OGGETTI_(query di estrazione elenco oggetti dei progetti)
	schema


		SELECT 
				"Elenco oggetti" AS Tx01, 
				PROGETTI_Msys_TB03_ELENCO_ObjProgetti.*, "" AS TxtChiavi, PROGETTI_Msys_TB03_ELENCO_ObjProgetti.COD_PROGETTO_s AS KeyCOD_PROGETTO_s, PROGETTI_Msys_TB03_ELENCO_ObjProgetti.TIPOGGETTO_s AS KeyTIPOGGETTO_s, PROGETTI_Msys_TB03_ELENCO_ObjProgetti.COD_PROGETTO_s AS ORD_COD_PROGETTO, PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NRO_OGGETTO_i AS ORD_NRO_OGGETTO_i
			FROM 
				PROGETTI_Msys_TB03_ELENCO_ObjProgetti;
