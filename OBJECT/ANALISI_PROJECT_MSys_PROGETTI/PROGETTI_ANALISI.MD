# PROGETTI_ANALISI.MD
### Note:  analisi sul micro progetto DENOMINATO @PROGETTI
@ANALISI@PROGETTO_(di gestione dei progetti nel db denominato MSys_PROGETTI_GE.mdb)

   QUOTA_PROGETTO
      Progetti è un micro sistema di definizione dei progetti che vengono gestiti in msys.
      Viene definito un CODICE PROGETTO con la descrizione ed un QUOTA PROGETTO che servira
      sulla tabella collegata PER DEFINIRE I GRUPPI DI OGGETTI e successivamente il nome di
      ogni singolo oggetto. In questo modo non ci saranno nomi duplicati per gli oggetti
      QUERY, TABLE, FORM REPORT MODULI ECC.. e ci saranno codice di raggruppamento che renderanno
      piu facile la estrazione dei gruppi di oggetti da creare e/o da modificare o manutenere.


   ATTIVO FILTRO PAGINE DA 1-12
      Note
         nella form di gestione dei progetti: PROGETTI_Msys_Tb00FrmM01_GE_PROGETTI
         ho attiviato il filtro della casella di testo TXT_FILTRO, con la quale eseguo
         2 operazioni:
            1) DOPPIO CLICK DEL CAMPO COD_PROGETTO_s 
               della sottoform_01 salvo il codice progetto nella casella di testo 
               TXT_FILTRO, vedi codice ---> TXT_FILTRO_DOPPIO_CLICK

            2) ATTIVA O DISATTIVA I FILTRI delle pagine da 1-8
               vedi codice: codice ---> @ATTIVA+DISATTIVA@FILTRI@PAGINE.DA_1-8


   QUERY DI ESTRAZIONE DEL GRUPPO CODICI

      Sulla tabella PROGETTI_Msys_TB02_TIPO_OGGETTI viene creata una query per estrarre il GRUPPO del codice QUOTE .

      La query PROGETTI_Msys_TB03Qry01_03_GROUP_QUOTE_s 




   OGGETTI DELLA GESTIONE DEI PROGETTI
      Note
         Gli oggetti per la gestione dei progetti sono :


         LE TABELLE

            MSys_DF02_TIPO_OGGETTO
            MSys_DF11_}----------------------------------------------------@
            Msys_DF11_PROGETTI
            PROGETTI_MSYS_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab
            PROGETTI_Msys_TB01_}-------------------------------------------@
            PROGETTI_Msys_TB01_PROJECT
            PROGETTI_Msys_TB01_PROJECT_TMP
            PROGETTI_Msys_TB02_}-------------------------------------------@
            PROGETTI_Msys_TB02_TIPO_OGGETTI
            PROGETTI_Msys_TB03_}-------------------------------------------@
            PROGETTI_Msys_TB03_ELENCO_ObjProgetti

         LE QUERY
            