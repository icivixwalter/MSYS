# README_MINI_PROGETTO_IMPORTA_ESPORTA_OGGETTI


### MODIFICHE
   2025_10_19 = ATTIVATO HELP
         Ho attivato l'help sia pdf che project. 
            vedi codice---> @2025_10_19@MODIFICA@ATTIVATO@HELP@PROGETTI

   2025_10_19 = CORRETTO ERRORE NULL 
         nel doppio click del progetto che serve per aggiornare con il codice la tabella tmp ho
         corretto l'errore del campo testo null (COD_PROGETTO_s_TXT_DblClick). Se è valorizzato salva in tmp il codice altrimenti salta
         la routine.
            vedi codice---> @2025_10_19@MODIFICA@COD_PROGETTO@SALVATA@TMP@SE@>NULL

### GE_MSys_ESPORTA_IMPORTA.mdb
      NOTE
         CODICE----> @PROGETTO@ESPORTA@IMPORTA@OGGETTI
         il progetto si trova in questa path ed è costruito per le esportazioni e le importazioni:
            start "ESPORTA IMPORTA" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_ESPORTA_IMPORTA\"

         Il db viene utilizzato per la gestione di tutti gli oggetti.

      CARATTERISTICHE IMPORTAZIONI ED ESPORTAZIONI
         Per le esportazioni e le importazioni occorre seguire le seguenti regole:

            A) IMPORTAZIONE NEL DB CORRENTE
               prevede di importare nel db corrente gli oggetti da db esterno con il vincolo che non si
               puo importare dal db corrente.

            C) ESPORTAZIONE DAL DB CORRENTE  
               Dal db corrente dove si trova il gestore puoi esportare in db esterni tranne su stesso.
            D) IMPORTI O ESPORTI DA DB DIVERSI
               dalla posizione neutra puoi attivare le importazioni dal db scelto oppure esportare da
               un db ad altro purche diversi.

      

      MODIFICHE
         ripristinato i campi COD_PROGETTO_s e PROGETTO_s nelle tabelle
            Msys_ESPORTA_IMPORTA_DB_DEST
            Msys_ESPORTA_IMPORTA_ELENCO_OGGETTI_DB
      
      IMPLEMENTAZIONI DELLE ATTIVITA
         1) IMPLEMENTAZIONE TABELLE 
            occorre implementare le tabelle dei progetti con quella imposta esporta oggetti in modo
            da avere lo stesso tipo di campi.
                  PROGETTI_Msys_TB03_ELENCO_ObjProgetti
                           |
                           |
                  Msys_ESPORTA_IMPORTA_ELENCO_OGGETTI_DB

            Per trovare l'esatta corrispondenza tra le due tabelle costruisco due query
            parallele per individuare i campi se sono uguali:
               CONFRONTA_Qry01_ESPORTA_IMPORTA
               CONFRONTA_Qry01_TB03_ELENCO_ObjProgetti
CHATGPT 
    TODO: da fare...
