MSys_TABELLE_LINK.md
	Note
		tabella msys da utilizare per il link dei progetti
		Per il link delle tabelle per ora sono state create:

	OGGETTI
		TABELLA
			MSys_TABELLE_LINK  = per il salvataggio della path del db e della tabella

		QUERY
			MSys_TABELLE_LINK_Qry_}----------------------------------------@
			MSys_TABELLE_LINK_Qry01_01_SELECT_TUTTE

		MODULO
			UTILITA_Msys_Mdl05_COLLEGA_TABELLE_LINK**




	ATTIVITA
		gli oggetti da collegare sono identificati condi il seguente
		CODICE ---> @MSys_TABELLE_LINK_(@tabelle@link per il collegamento)



	GIORNALE_CORREZIONI
        2024.08.07 = 






    



    FUNZIONI DI RICERCA TESTATE NEL MODULO UTIL_MDL40_N09_CONTROLLO_ESISTENZA_OGGETTO_DB

        Controllo se esiste la form     :  pfCONTROLLO_ESISTENZA_OGGETTI_Form
                                        '//@CONTROLLO@ESISTENZA@FORM_(Funzione che rileva se esiste nel db corrente la form restituendo True o False)


        Controllo se esiste il REPORT   :  pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS
                                        '//@CONTROLLO@ESISTENZA@REPORT_(Funzione che rileva se esiste nel db corrente la REPORT restituendo True o False)



        Controllo se esiste IL MODULO  :  pfCONTROLLO_ESISTENZA_OGGETTI_MODULI
                                '//@CONTROLLO@ESISTENZA@MODULI_(Funzione che rileva se esiste nel db corrente dei MODULI o del @MODULO restituendo True o False)




        Controllo se esiste DELLA MACRO  :  pfCONTROLLO_ESISTENZA_OGGETTI_MACRO
                                '//@CONTROLLO@ESISTENZA@MACRO_(Funzione che rileva se esiste nel db corrente dei MACRO restituendo True o False)


