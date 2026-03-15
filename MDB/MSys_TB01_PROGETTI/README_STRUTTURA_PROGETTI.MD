# STRUTTURA DEI PROGETTI
## FILE
il file per lagestione della struttura dei progetti:
README_STRUTTURA_PROGETTI.MD


2026_03_15_ESPORTA_PROGETTO = definisci la procedura di spostamento
degli oggetti del progetto , dal database 
   MSys_TB04_DB_ESTERNO.mdb  ----> GE_MSys_TB04_DB_ESTERNO.mdb
   ogetti:
      query
         PROGETTI_Msys_TB03Qry00_{@========ELENCO_ObjProgetti==========@}
         PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti
         PROGETTI_Msys_TB04Qry01_}--------------------------------------@
         PROGETTI_Msys_TB04Qry01_01_Database_DEST

      le form 
         ....


SCHEMA DEFINIZIONE PROGETTI
         Nel mini progetto DF11 definisci solo il codice la denominazione e le quote di ogni codice progetto, quindi solo
         i codici essenziali. Dai codici progetto puoi sviluppare il vero progetto ed i suoi oggetti utilizzando le tabelle
         TB01, TB02 E TB03 del progetto denominato Msys_TB01_PROGETTI.mdb. La struttura delle tabelle master e quelle
         detail è la seguente:
      Msys_DF11_PROGETTI                                    = Tabella di definizione dei codice di progetto
            
        ## FUNZIONAMENTO
            ad ogni codice progetto corrisponde un PROGETTO della DF11 che viene inserito nella tabella
            PROGETTI_Msys_TB01_PROJECT CON LA DESCRIZONE DEL PROGETTO. Poi la seconda tabella individua
            la classificazione di oggetti da aggiungere : QUERY, TABLE MODULI ECC.
            La terza tabella definsce per ogni tipo di oggetto il suo nome mentre la QUARTA TABELLA 
            stabilisce per ogni singolo oggetto il db di destinazione.
      ## STRUTTURA LOGICA
         PROGETTO
            → TIPO OGGETTO
                 → OGGETTO
                      → DATABASE DESTINAZIONE
                            
      ## STRUTTURA MASTER/DETAIL
      |
      |
      |-----> PROGETTI_Msys_TB01_PROJECT            = TABELLA 1 = Progetti (codice e descrizione)
               |
               | key master/detail = COD_PROGETTO_s
               |
              PROGETTI_Msys_TB02_TIPO_OGGETTI       = TABELLA 2 = Tipi oggetti del progetto
               |                                      (QUERY, TABELLE, FORM, REPORT...)
               |
               | key master/detail = COD_PROGETTO_s ; TIPOGGETTO_s
               |
              PROGETTI_Msys_TB03_ELENCO_ObjProgetti = TABELLA 3 = Elenco oggetti del progetto
               |                                      necessari al funzionamento
               | key master/detail = COD_PROGETTO_s ; TIPOGGETTO_s ; NOMEOGGETTO_s
               |
              PROGETTI_MSYS_TB04_DB_ESTERNO         = TABELLA 4 = Database esterno di destinazione
                                                      dell'oggetto

