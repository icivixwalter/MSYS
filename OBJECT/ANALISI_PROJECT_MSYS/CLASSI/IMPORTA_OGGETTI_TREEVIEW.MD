IMPORTA_OGGETTI_TREEVIEW.MD
	


GIORNALE_AGGIORNAMENTI
	CMB_02
		2024
			2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODICA
				nella selezione oggetti della combinata 02 vi è 
				l'errore di non impostare la path della casella
				di testo sourceDBPath_s_Txt. Esempio se seleziono
				IMPORTA_OGGETTI_MSYS_DLL la path è correttamente:
					c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb
				se seleziono 
				IMPORTA_OGGETTI_TREEVIEW, path è errata in quanto lascia
					c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_N01_OGGETTI.mdb
				mentre dovrebbessere quella della treeView.

				L'errore è dovuto perche la combinata due permette la selezione degli oggetti MSYS
				e TreeView. L'errore è logico, perchè la combinata 01 che imposta la path per tutti
				i progetti.

				RISOLUZIONE: La soluzione si trova nella cmb_02 che deve filtrare nelle voci della cmb_02
					solo oggetti Msys non altro.

