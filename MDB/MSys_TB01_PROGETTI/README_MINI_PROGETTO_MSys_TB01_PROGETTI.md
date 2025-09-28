# README_MINI_PROGETTO_Msys_TB01_PROGETTI

### GE_Msys_TB01_PROGETTI.mdb
   NOTE
      progetto DI GESTIONE DI TUTTI PROGETTI che viene eseguito mediante il PROGETTO
      ACCCESS: 
         GE_Msys_TB01_PROGETTI.mdb
      In esso vengono gestite :
         a) la definizione del codice progetto con l'utilizzo deglio oggetti del mini progetto
            DF11.
               FORM
                  Msys_DF11_Frm01_M01_GE_CODICE_PROGETTI_TUTTI
                     si trova in questa path:
   
                        start "apri PROGETTI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB01_PROGETTI\MDB\GE_Msys_TB01_PROGETTI.mdb"

         B) i PROGETTI 
            ossia tutti gli oggetti per gestire il singolo MINI PROGETTO in 
               start "apri gestione oggetti dei MINI PROGETTI" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB01_PROGETTI\MDB\GE_Msys_TB01_PROGETTI.mdb"
   CORREZIONE FORM --> PROGETTI_Msys_TB02Frm01_01_TIPO_OGGETTI
               NEL DB : GE_Msys_TB01_PROGETTI.mdb
                           |
                           | --->FORM -> PROGETTI_Msys_TB02Frm01_01_TIPO_OGGETTI
                                                |
                                                |__Non funziona nel campo ---> TIPO_OGGETTO = risolto
   CORREZIONE FORM --> PROGETTI_Msys_TB03Frm01_01_ELENCO_ObjProgetti
               NEL DB : GE_Msys_TB01_PROGETTI.mdb
                              |
                              |____________FORM
                                           |
                                           |__PROGETTI_Msys_TB03Frm01_01_ELENCO_ObjProgetti
                                                      |  
                                                      |__COD_PROGETTO = casella di controllo non funziona.
                                                                        risolto
   CORREZIONE FORM --> Msys_DF11_Frm01_S01_COD_PROGETTI
               NEL DB : GE_Msys_TB01_PROGETTI.mdb
                              |
                              |____________FORM
                                           |
                                           |__Msys_DF11_Frm01_S01_COD_PROGETTI
                                              |    
                                              |-----PATH_DEST_s = casella di controllo ammessa a tutti gli elenchi
   CORREZIONE FORM --> Msys_DF01_Frm01_M01_GE_COMANDI
               NEL DB : PATH apri con start "APRI PROGETTO COMANDI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\GE_MSys_DF01.mdb"
                              |
                              |____________FORM
                                           |
                                           |__Msys_DF01_Frm01_M01_GE_COMANDI
                                              |    
                                              |-----COD_PROGETTO_s_txt     = casella di controllo impostata query
                                              |-----PROGETTO_s_TXT         = casella di controllo impostata query
                                              |-----COMANDO_s_TXT          = casella di controllo impostata query
                                              |-----per tutte le caselle combinate: SELECT MSys_DF01_COMANDI.COMANDO_s FROM MSys_DF01_COMANDI; 
### IMPLEMENTAZIONE MSYS_PROGETTI CON MSYS_ESPORTA
   Note
      GE_Msys_TB01_PROGETTI.mdb
         Con il progetto TB01_PROGETTI  eseguo due attivita:
         01) Msys_DF11_Frm01_M01_GE_CODICE_PROGETTI_TUTTI   = gestisco i codici dei progetti
               con il campo COD_PROGETTO_s_txt, attivo la creazione e cancellazione dei
               codici di progetto
         02) PROGETTI_Msys_TB01Frm00_GE_PROGETTI            = gestisco gli oggetti dei progetti   
               cioè tutti gli oggetti necessari per il funzionamento del progetto, di cui 
               query
               table
               form
               report
               moduli
   ESPORTA OGGETTI CON Msys_Frm01_S01_ESPORTA_DB_EST_00_Titolo
      Note
         con il progetto del database GE_MSys_ESPORTA.mdb esporto gli oggetti dal db
         corrente ESPORTA per ora in un database di destinazione e si basa sulla tabella
         Msys_ESPORTA_ELENCO_OGGETTI_DB
         che contiene tutti gli oggetti.
      IMPLEMENTAZIONE
         si dovrebbe implementare le tabella del progetto TB01 con il progetto ESPORTA
         ma come:
            a) una ipotesi è quella di salvare gli oggetti da TB01 ----> ESPORTA
               vantaggi:
                  - rimane l'atonimicita del progetto con tabella e gestione;
                  - todo: si deve distinguere meglio la funzionalita delle tabelle
                        Msys_ESPORTA_DB_EST              = questa tabella rappresenta la destinazione
                        Msys_ESPORTA_ELENCO_OGGETTI_DB   = questa tabella rappresenta gli oggetti da importare
                                                           o esportare e quindi ogni singolo oggetto ha il db di origine

                 - todo : occorre in questo caso effettuare il controllo prima della IMPORTAZIONE o della esportazione:
                          01)  IMPORTAZIONE DB CORRENTE
                                 In questo primo caso in cui il progetto ESPORTA VIENE UTILIZZATO PER L'IMPORTAZIONE per cui
                                 deve essere effettuato il confrontonto tra le due tabelle
                                    Msys_ESPORTA_DB_EST = Msys_ESPORTA_ELENCO_OGGETTI_DB
                                 se le path e il db delle due tabelle coincide non puo essere effettuata la importazione 
                                 tra il db corrente che importa ed i singoli oggetti da importare i quali vengono presi dal
                                 db corrente e quindi importati su se stesso.
                                 E' ammesso solo l'importazione nel db corrente con OGGETTI CHE PROVENGON DA UB DB CON ORIGINE DIVERSA 
                          02) ESPORTAZIONE DAL DB CORRENTE
                              NON E' AMMESSA L'ESPORTAZIONE SU SE STESSO   per cui il db di destinazione <> da origine

                          03) ESPORTA DA ALTRO DB
                              il db di partenza è diverso dagli oggetti che hanno come db origine diverso da quello
                              di partenza per cui gli oggetti partono dal Db origine ---> DB DESTINAZIONE

                           04) IMPORTA TRA DB DIVERSI
                              il db di partenza è diverso dai db degli oggetti di origine.


                  todo: rinominare le tabelle e il mini progetto in MSYS_ESPORTA_IMPORTA_OGG
                        per evitare di creare tabelle di esportazioni e di importazioni separate con duplicazione dei dati
                        occorre modificare il NOME DI TUTTI GLI OGGETTI DEL DATABASE in ESPORTA IMPORTA OGGETTI DATABASE
                        Il PROGETTO SI DEVE RINOMINARE IN 
                              GE_MSys_ESPORTA_IMPORTA.mdb
                        TABELLE
                              Msys_ESPORTA_IMPORTA_OGG               = in questa tabella il database di DESTINAZIONE
                              Msys_ESPORTA_IMPORTA_ELENCO_OGG        = ELENCO DEGLI OGGETTI SORGENTE ed i SINGOLI DB DI ORIGINE




            b) seconda ipotesi è quella di sostituire ESPORTA da TB01.
               svantaggi: si perdo l'atonimicita del progetto








CHATGPT 
    TODO: da fare...

### RIDENOMINAZIONE DEL MINI PROGETTO DA GE_MSys_ESPORTA.mdb ---> GE_MSys_ESPORTA_IMPORTA.mdb
   CREATO CARTELLA MSys_ESPORTA_IMPORTA
      la nuova cartella rappresenta il mini progetto per la gestione delle ESPORTAZIONE degli oggetti E DELLE IMPORTAZIONI
      che puo avvenire tra:
         - db corrente e db di origine
         - db esterni diversi
   CREATO MINI PROGETTI GESTIONE
      i mini progetti creati sono:
            GE_MSys_ESPORTA_IMPORTA             = progetto master di gestione esporta ed importa oggetti

   CREATO MINI PROGETTO TABELLE
      MSys_ESPORTA_IMPORTA.mdb                  = PROGETTO DELLE TABELLE

         TABELLE
            Msys_ESPORTA_IMPORTA_DB_DEST                    = TABELLA PROGETTO DI DESTINAZIONE
            Msys_ESPORTA_IMPORTA_ELENCO_OGGETTI_DB          = TABELLA PROGETTI ED OGGETTI DI ORIGINE trattati

         QUERY ****  CONTINUARE ****
