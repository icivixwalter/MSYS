@query@name@MSys_TB0401Qry01_01_GE_TabContr_MASTER_(la query di estrazione dei progetti gestiti DAL TAB CONTROL MASTER)

	Note
			@PROGETTO@TAB@CONTROL@MASTER_(i progetti che utilizzano il tab control master, vedi codice tb0401)

		il TAB CONTROL MASTER gestisce 12 pagine in cui caricare le 12
		form e si applica alla FORM GENERALE MASTER.
		Si basa sulla seguente query che si basa sulla tabella
		MSys_TB0401_GE_TabContr_MASTER:


	schema

		SELECT 
			MSys_TB0401_GE_TabContr_MASTER.*, 
			"->" AS TxtChiavi, 
			MSys_TB0401_GE_TabContr_MASTER.COD_PROGETTO_s AS KeyCOD_PROGETTO_s
			FROM 
			MSys_TB0401_GE_TabContr_MASTER;




	tabella_tmp
		vi è anche la query tmp che si basa sulla tabella temporanea denominata
		@query@name@MSys_TB0401Qry01_01_GE_TabContr_MASTER_TMP_(la query di estrazione delle pagine del TAB CONTROL MASTER)



	schema
		SELECT 
			MSys_TB0401_GE_TabContr_MASTER_TMP.*, 
			"->" AS TxtChiavi, 
			MSys_TB0401_GE_TabContr_MASTER_TMP.COD_PROGETTO_s AS KeyCOD_PROGETTO_s
			FROM 
			MSys_TB0401_GE_TabContr_MASTER_TMP;





	RIEPILOGO_NOME_QUERY
		MSys_TB0401Qry00_}---------------------------------------------@
		MSys_TB0401Qry01_01_GE_TabContr_MASTER
			questa query è associata alla form 
					@form@associata_(MSys_TB0401Frm01_01_GE_TabContr_MASTER
						vedi codice---> @form@associata@MSys_TB0401Frm01 )

		MSys_TB0401Qry01_01_GE_TabContr_MASTER_TMP
