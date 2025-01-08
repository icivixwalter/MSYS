CODICE_PROGETTO_DblClick

CODICE_PROGETTO_EVENTO_

	Note
		evento doppioclick per i campo @COD_PROGETTO_s_Txt in cui attivo il 
		caricamento del progetto nelle sottoform



		qui salvo il codice del progetto da caricare:
		 	MSys_TABELLE_Tmp




		 	@IMPOSTA@FILTRO@OGGETTI@DA@CARICARE_(in questo parte del codice imposto il filtro degli oggetti da caricare)
		 		MA SOLO IN VISUALIZZAZIONE


		 	faq
		 		@dove_(vengono filtrati gli oggetti da caricare?)
		 			la query sql con il parametro che filtra gli oggetti da caricare si 
		 			trova con questo codice : 
		 				@TABELLA.CARICA.PAGINE
		 				@SQL@FILTRO@OGGETTI
	 				all'interno del progetto msys. Ma si tratta solo del FILTRAGGIO PARAMETRI
	 				non la query che viene inserita nelle proprieta della sottoform.


	 				@SQL@FILTRO@OGGETTI@PER@PARAMETRI_(Qui vengono filtrati i parametri da inviare alla funziona imposta form)

	 					Per aprire il filtro in formato pdf apri con il DOS

	 					START "APRI IL FILTRO SQL" "c:\CASA\CDM\GE_CDM\MSYS\MSYS_ANALISI\Msys_PROGETTO\PDF\FILTRO_SQL.pdf"
