
@QUERY@ESPORTA_(qui @TUTTE@LE@query@esporta database, select, delete insert ed update )

SELECT
	Msys_Qry01_01_ESPORTA_DB_EST.MD
			Note
				query per la gestione del database
					@query@esporta_(la query esporta dei @database)


			SCHEMA

				SELECT 
					Msys_ESPORTA_DB_EST.COD_PROGETTO_s, 
					Msys_ESPORTA_DB_EST.PATH_s, 
					Msys_ESPORTA_DB_EST.DATABASE_s, 
					Msys_ESPORTA_DB_EST.NOTE_s
					FROM 
						Msys_ESPORTA_DB_EST
					WITH 
						OWNERACCESS OPTION;

			Msys_Qry01_01_ESPORTA_DB_EST_TMP
			Note
				la tabella temporanea di esportazione del db esterno per i salvataggi temporanei.
					@query@select@tmp_(la query di estrazione dei dati dalla @tabella@tmp)

	SCHEMA_TABELLA_TMP

			SELECT 
					Msys_ESPORTA_DB_EST.Scel_b, 
					Msys_ESPORTA_DB_EST.COD_PROGETTO_s, 
					Msys_ESPORTA_DB_EST.PATH_s, 
					Msys_ESPORTA_DB_EST.DATABASE_s, 
					Msys_ESPORTA_DB_EST.NOTE_s
				FROM 
					Msys_ESPORTA_DB_EST
				WITH OWNERACCESS OPTION;

DELETE
		
	Msys_Qry01_01_ESPORTA_DB_EST_TMP_Delete
		Note
			la cancellazione dei dati nella tabella tmp
				@query@delete_(la query di @cancellazione dei @dati@tmp)
		SCHEMA

			DELETE 
				Msys_ESPORTA_DB_EST_TMP.*
				FROM 
				Msys_ESPORTA_DB_EST_TMP
				WITH 
				OWNERACCESS OPTION;




INSERT
	Msys_Qry01_01_ESPORTA_DB_EST_Insert_To_Tmp
		Query di inserimento dei dati in tabella tmp

		INSERT INTO 
			Msys_ESPORTA_DB_EST_TMP 
				( COD_PROGETTO_s, PATH_s, DATABASE_s, NOTE_s )
			SELECT 
				Msys_ESPORTA_DB_EST.COD_PROGETTO_s, 
				Msys_ESPORTA_DB_EST.PATH_s, 
				Msys_ESPORTA_DB_EST.DATABASE_s, 
				Msys_ESPORTA_DB_EST.NOTE_s
			FROM 
				Msys_ESPORTA_DB_EST
			WITH 
				OWNERACCESS OPTION;



		STRINGA_SSQL
			Note
				2024_08_18=EventoDoppioclick_CampoCOD_PROGETTO_s_Txt_Sql
					Stringa ssql per l'inserimento in tabella TMP de codice progetto e della path + Note

				sSql=sSql & "INSERT INTO "
				sSql=sSql & "Msys_ESPORTA_DB_EST_TMP "
				sSql=sSql & "( COD_PROGETTO_s, PATH_s, DATABASE_s, NOTE_s )"
				sSql=sSql & "SELECT "
				sSql=sSql & "Msys_ESPORTA_DB_EST.COD_PROGETTO_s, "
				sSql=sSql & "Msys_ESPORTA_DB_EST.PATH_s, "
				sSql=sSql & "Msys_ESPORTA_DB_EST.DATABASE_s, "
				sSql=sSql & "Msys_ESPORTA_DB_EST.NOTE_s"
				sSql=sSql & "FROM "
				sSql=sSql & "Msys_ESPORTA_DB_EST"
				sSql=sSql & "WITH "
				sSql=sSql & "OWNERACCESS OPTION;"
