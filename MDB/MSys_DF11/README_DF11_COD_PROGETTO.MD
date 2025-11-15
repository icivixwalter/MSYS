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

SCHEMA DEFINIZIONE PROGETTI
         Nel mini progetto DF11 definisci solo il codice la denominazione e le quote di ogni codice progetto, quindi solo
         i codici essenziali. Dai codici progetto puoi sviluppare il vero progetto ed i suoi oggetti utilizzando le tabelle
         TB01, TB02 E TB03 del progetto denominato Msys_TB01_PROGETTI.mdb. La struttura delle tabelle master e quelle
         detail è la seguente:
      Msys_DF11_PROGETTI                                    = Tabella di definizione dei codice di progetto
            |
            |-----> PROGETTI_Msys_TB01_PROJECT              = Tabella in cui vengono gestiti gli oggetti del singolo progetto
                     |
                     |
                    PROGETTI_Msys_TB02_TIPO_OGGETTI         = La tb01 è colleata alla tb02 per definire i gruppi di oggetti gestiti
                     |                                      es. tutte le query, form, report ecc...
                     |
                    PROGETTI_Msys_TB03_ELENCO_ObjProgetti   = la TB02 è collegata a TB03 dove sono salvati tutti gli oggetti necessari
                                                               per il funziomento del progetto.
                                                            Tutti i collegamenti della sotto tabelle si trovano i questa path:
                                                            START "APRI PROGETTI" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB01_PROGETTI\MDB\Msys_TB01_PROGETTI.mdb"

2025.11.15 = SPECIFICATO IL RUOLO DEL CODICE PROGETTO
          il database contiene la tabella in cui vengono specificati i codici di progetto con
          la descrizione e la quota di progetto che rappresenta un secondo codice in cui filtrare
          gli oggetti specifici del progetto es.
          Nel progetto ANAGRAFE abbiamo il codice progetto definito in CDM_ANAGRAFE descrizione
          ed altro. Il codice quota del progetto anagrafe è la seguente:
               GROUP_QUOTA_CDM_ANAGRAFE
          con questo codice puoi identificare gli oggetti che appartengono al progetto divisi
          e distinti da questo codice es:
            QUERY ANAGRAFE          = GROUP_QUOTA_CDM_ANAGRAFE_Query 
                                       identifico in modo unico tutte le query che appartengono
                                       al progetto anagrafe.
            QUERY ANAGRAFE          = GROUP_QUOTA_CDM_ANAGRAFE_Form 
                                       anche qui identifico tutte le form del progetto medesimo.
         con i codice quota puoi indentificare tutti gli oggetti ed assoggettarl alle lavorazioni necessarie
         in modo veloce.
         Il problema da risolvere è in quale progetto .mdb si possono definire gli oggetti di lavorazione.
         Una soluzione è stata trovata eliminando Msys_DF11_PROGETTI che non ha collegamenti utili con 
         lo sviluppo dei vari progetti gia avviati e sostituita con il link PROGETTI_Msys_TB01_PROJECT del 
         file Msys_TB01_PROGETTI.mdb. Questo mini progetto gestisce tutti gli oggetti occorenti per il loro
         funzionamento.

