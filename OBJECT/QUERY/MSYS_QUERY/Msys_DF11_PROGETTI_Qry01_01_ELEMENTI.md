Msys_DF11_PROGETTI_Qry01_01_ELEMENTI.md
	Note
	query adibita esclusivamente alla definizione dei progetti
	vedi codice 2024_08_23_DF11_PROGETTI_RIDEFINITO





	schema

			SELECT 
					Msys_DF11_PROGETTI.COD_PROGETTO_s, 
					Msys_DF11_PROGETTI.PROGETTO_s, 
					Msys_DF11_PROGETTI.NOTE_PROGETTO_s, 
					Msys_DF11_PROGETTI.DISCO_ORIG_s, 
					Msys_DF11_PROGETTI.PATH_ORIG_s, 
					Msys_DF11_PROGETTI.DATABASE_ORIG_s, 
					Msys_DF11_PROGETTI.NOTA_PROG_ORIG_s, 
					Msys_DF11_PROGETTI.DISCO_DEST_s, 
					Msys_DF11_PROGETTI.PATH_DEST_s, 
					Msys_DF11_PROGETTI.DATABASE_DEST_s, 
					"le chiavi ---->" AS TxtChiavi, 
					Msys_DF11_PROGETTI.COD_PROGETTO_s AS KeyCOD_PROGETTO_s
				FROM 
				Msys_DF11_PROGETTI;
