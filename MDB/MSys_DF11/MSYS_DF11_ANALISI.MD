# SUB PROGETTO MSYS_DF11_ANALISI.MD

## NOTE
	nella tabella @DF11 vengono salvati i @CODICI@PROGETTO e la loro definizione.
		@analisi@DF11_(analisi per msys DF11)

		
			La path della gestione DI DF11:
				@apri@path@MSYS_df11_(@path@DF11 si trova in questa directory detta
					anche @df11)

				start "apri TOTAL DF11" "C:\Program Files\totalcmd\TOTALCMD64.EXE" left: c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF11\, right: c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF11\ 
	

	
	MODIFICHE
		DB  MSys_DF11.mdb
				|
				|__FORM ---> Msys_DF11_Frm01_S01_COD_PROGETTI
							  			|
							  			|__CAMPO ----> COD_PROGETTO_s_txt
							  			|				   INSERITO QUERY: SELECT Msys_DF11_PROGETTI.COD_PROGETTO_s FROM Msys_DF11_PROGETTI GROUP BY Msys_DF11_PROGETTI.COD_PROGETTO_s ORDER BY Msys_DF11_PROGETTI.COD_PROGETTO_s WITH OWNERACCESS OPTION;
							  			|
							  			|__CAMPO ----> COD_PROGETTO_s_txt
							  			|				   inserito query: SELECT Msys_DF11_PROGETTI.PROGETTO_s, Msys_DF11_PROGETTI.COD_PROGETTO_s FROM Msys_DF11_PROGETTI GROUP BY Msys_DF11_PROGETTI.PROGETTO_s, Msys_DF11_PROGETTI.COD_PROGETTO_s ORDER BY Msys_DF11_PROGETTI.COD_PROGETTO_s WITH OWNERACCESS OPTION;
							  			|__CAMPO ----> DISCO_ORIG_s_TXT
							  			|					inserito query:	SELECT [MSys_DF10_DATABASE].[DISCO_s] FROM MSys_DF10_DATABASE GROUP BY [MSys_DF10_DATABASE].[DISCO_s] ORDER BY [MSys_DF10_DATABASE].[DISCO_s] WITH OWNERACCESS OPTION; 
							  			|__CAMPO ----> PATH_ORIG_s_TXT
							  			|					inserito query:	SELECT MSys_DF10_DATABASE.PATH_s, MSys_DF10_DATABASE.DISCO_s FROM MSys_DF10_DATABASE GROUP BY MSys_DF10_DATABASE.PATH_s, MSys_DF10_DATABASE.DISCO_s ORDER BY MSys_DF10_DATABASE.PATH_s WITH OWNERACCESS OPTION;

							  			|
							  			|
							  			|
							  			|		

1	
		


2
		

3
		

