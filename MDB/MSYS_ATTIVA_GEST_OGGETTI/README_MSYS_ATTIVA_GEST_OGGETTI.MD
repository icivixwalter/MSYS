# READM ATTIVA GESTIONE OGGETTI

# EVENTO ATTIVA PROGETTI
   COD_PROGETTO_s_TXT_DblClick(Cancel As Integer)        = ATTIVO IL CODICE PROGETTO SCELTO
      Qui viene fatto il controllo se la form master è aperta altrimenti no ossia:
         AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER deve essere aperta







# STRUTTURA DEI PROGETTI

## FILE

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

