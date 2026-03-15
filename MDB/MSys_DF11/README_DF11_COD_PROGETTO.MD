# @CODICE@PROGETTO
   NOTE
      gestione del file  README_DF11_COD_PROGETTO.MD per la definizione e gestione dei
      CODICI DI PROGETTO. Per ogni codice progetto occorre inserire:
         COD_PROGETTO_s                = il codice progetto
         CDM_ANAGRAFE

         PROGETTO_s                    = il nome del progetto es.
         PROGETTO_ANAGRAFE_CONDOMINIALE

         NOTE_PROGETTO_s               = note del progetto

         NRO_DEF_i                     = nro definzione progetto come indice
         1
         GROUP_QUOTE_s                 = gruppo quote per definire le quote di ogni oggetto
         GROUP_QUOTA_CDM_ANAGRAFE        ad es. GROUP_QUOTA_CDM_ANAGRAFE_frm = tutte le form
                                                GROUP_QUOTA_CDM_ANAGRAFE_Qry = tutte le query
                                                            che appartengono al progetto
         DESCR_QUOTE_S                 = descrizione della quota
         QUOTA CONDOMINIO GESTIONE  ANAGRAFE
      Tutti i progetti sono definiti come oggetti nel db Msys_TB01_PROGETTI.mdb
      con la tabella PROGETTI_Msys_TB01_PROJECT
