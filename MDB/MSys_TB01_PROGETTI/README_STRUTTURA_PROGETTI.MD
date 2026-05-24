# STRUTTURA DEI PROGETTI
## FILE
il file per lagestione della struttura dei progetti:
README_STRUTTURA_PROGETTI.MD


2026_03_15_ESPORTA_PROGETTO *** da finire *** = definisci la procedura di spostamento
degli oggetti del progetto , dal database 
   MSys_TB04_DB_ESTERNO.mdb  ----> GE_MSys_TB04_DB_ESTERNO.mdb
   ogetti:
      query
         PROGETTI_Msys_TB03Qry00_{@========ELENCO_ObjProgetti==========@}
         PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti
         PROGETTI_Msys_TB04Qry01_}--------------------------------------@
         PROGETTI_Msys_TB04Qry01_01_Database_DEST

Nel database GE_Msys_TB01_PROGETTI.mdb ho modicato la form 
   PROGETTI_Msys_TB03Frm01_01_ELENCO_ObjProgetti  che deve essere salvata
   negli altri database.
   La modifica al fine di semplicare l'inserimento degli oggetti.

Nel database MSys_ESPORTA.mdb ho salvato il modulo
   Msys_GEFILE_Mdl02_ESPORTA_OGGETTI

   questo modulo deve essere modificato per eseguire le esportazioni 
   con queste tabelle:

   MSys_TB01_PROGETTI                  = qui deve recuperare il codice progetto
                                         quello di prova denominato MSys_TB01_PROGETTI
TODO: 01) creare una query che estragga il progetto con questo
      codice MSys_TB01_PROGETTI dalla tabella MSys_TB01_PROGETTI
      02) estrarre dal codice progetto il database di origine
            path di origine
            database di origine


TODO: dalla 2 tabella  collegatata PROGETTI_Msys_TB02_TIPO_OGGETTI
      estrarre con il codice progetto TB04_to_GE_TB04
         01) il valore del campo TIPOGGETTO_s
         02) il valore del campo COD_PROGETTO_s

TODO: dalla 3 tabella collegata PROGETTI_Msys_TB03_ELENCO_ObjProgetti con
      il valore del codice progetto TB04_to_GE_TB04 estrarre i 
      seguenti valori:
         01) il valore del codice progetto dal campo COD_PROGETTO_s
         02) il valore del tipo oggetto dal campo TIPOGGETTO_s
         03) il valore nome dell'oggetto query dal campo NOMEOGGETTO_s
         

TODO: dalla 4 tabella collegata PROGETTI_MSYS_TB04_DB_ESTERNO i 
     seguenti valori con il codice progetto TB04_to_GE_TB04:
         01) il valore del codice progetto dal campo COD_PROGETTO_s
         02) il valore del tipo oggetto dal campo TIPOGGETTO_s
         03) il valore il nome dell'oggetto dal campo NOMEOGGETTO_s
         04) il DATABASE DI DESTINAZIONE  dal campo DATABASE_DEST_s
         05) il DISCO   dal campo DISCO_DEST_s
         06) LA PATH DI DESTINAZIONE  dal campo PATH_DEST_s
         
         In conclusione per ricavare dalla 4 tabella il dabatase su
         cui trasferire l'oggetto basta filtrare dalla tabella 4
         con i filtri:
            codice progetto
            tipo di oggetto
         forse ma è da verifiare puo recuperare la 
            path di destinazione
            il database di destinazione.

TODO: 
     Creare una funzione con la quale si ha :
      01) apri il database di partenza
      02) individua l'oggetto da esportare
      03) se esiste l'oggetto query attiva la esportazione
      04) apri il database di destinazione
      05) se esiste l'oggetto query da esportare prima cancella altrimenti no
      06) ESPORTA LA QUERY dal DATABASE DI PARTENZA  to DATABASE DI DESTINAZIONE
      7)  con l'oggetto transfert
      8)  




2026_03_15_CODICE PROGETTO = SOSTITUZIONE DELLA TABELLA Msys_DF11_PROGETTI CON LA Msys_DF11_COD_PROGETTO
   Adesso i codice progetto deve essere ricavato dalla tabella 
      Msys_DF11_COD_PROGETTO
      


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
                            
  

# STRUTTURA PROGETTI  NUOVA
   Msys_TB01_PROGETTI                           = TABELLA 1 = Progetti (codice e descrizione)
               |
               | key master/detail ;:COD_PROGETTO_s
               |
              Msys_TB02_TIPO_OGGETTI            = TABELLA 2 = Tipi oggetti del progetto
               |                                      (QUERY, TABELLE, FORM, REPORT...)
               |
               | key master/detail : COD_PROGETTO_s ; TIPOGGETTO_s
               |
              Msys_TB03_ELENCO_ObjProgetti      = TABELLA 3 = Elenco oggetti del progetto
               |                                      necessari al funzionamento
               | key master/detail : COD_PROGETTO_s ; TIPOGGETTO_s ; NOMEOGGETTO_s
               |
              MSYS_TB04_DB_ESTERNO :            TABELLA 4 = Database esterno di destinazione
              


# STRUTTURA OLD ======================================================

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

