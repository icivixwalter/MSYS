MySys_AGGIORNAMENTI

	2022.06.26_AggiornoTabellaTmp
		MSys_TUTTI_PROG_FrmM01_01_PROGETTI
			Note
				Modificato l'evento doppio clici sul campo 
					DENOM_PROGETTO_s


TABELLE
	2022.06.26_CreatoTabelleTmp

		Note
			Creato la tabella tabelle tmp per salvare il codice progetto e la denominazione e
			nella tmp precisamente nei campi
				COD_PROGETTO_s
				DENOM_PROGETTO_s
		MSys_TABELLE_Tmp



QUERY
	2022.07.02 
		CREATA UNA QUERY SULLA TABELLA TMP  MSys_TABELLE_Tmp
			*******da inserire nel progetto oggetti**********
			MSys_TABELLE_TMP_Qry_}-----------------------------------------@
			MSys_TABELLE_TMP_Qry01_01_SELECT_TUTTE
				Note
					estra le tabelle tmp oggetto di lavorazione
					@INSERIRE.NEL.PROGETTO

				struttura



		CREATA NUNA NUOVA QUERY DI CANCELLAZIONE DATI TMP
			MSys_TABELLE_TMP_Qry01_80_}------------------------------@DELETE
			MSys_TABELLE_TMP_Qry01_81_DELETE_TUTTE
				Note
					cancella dati nella tabella oggetti Table tmp


				struttura

					DELETE 	MSys_TABELLE_tmp.*
					FROM 	MSys_TABELLE_tmp;


FORM
		
	2022.07.03
		Note modifica della form di caricamento dei proggeti che puo essere utilizzata per qualsiasi
		progetto senza che rimanga bloccata solo su MSys_FORM_M01_GESTIONE_OGGETTI_MSys. Le modifiche
		nel codice si trovano su questo codice: @2022.03.07.modifiche

		MSys_TUTTI_PROG_FrmM01_01_PROGETTI