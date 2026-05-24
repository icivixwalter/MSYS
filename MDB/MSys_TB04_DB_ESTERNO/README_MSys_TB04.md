# MSys_TB04
   progetto DA ESPORTARE IN UN DATABASE ESTERNO nel file:
         MSys_TB04_DB_ESTERNO.mdb
   Che si trova in questa path:
         c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB04_DB_ESTERNO\MDB\

         START "apri db TB04 DB ESTERNO" CALL "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB04_DB_ESTERNO\MDB\MSys_TB04_DB_ESTERNO.mdb" ^& EXIT

   Questo progetto deve consentire l'esportazione degli oggetti nel database scelto
   nei campi DELLA TABELLA 
      PROGETTI_MSYS_TB04_DB_ESTERNO         :
         COD_PROGETTO_s                         = CODICE PROGETTO
         TIPOGGETTO_s                           = IL TIPO DI OGGETTO (QUERY , TABLE FORM ECC)
         NOMEOGGETTO_s                          = IL NOME DELL'OGGETTO
      questi 3 campi sono la chiave dellat abella DB_ESTERNO e contiene le path ed i database
      su cui esportare gli oggetti della tabella 3:
         PROGETTI_Msys_TB03_ELENCO_ObjProgetti

   todo: creare una query di prova che colleghi 3 database
      CODICE---> TB01



## TABELLE
La tabella DB_ESTERNO è l'utima tabella per la gestione dei progetti e contiene
la path ed il db dove salvare gli oggetti; questa tabella permette in questo modo di
stabibiere piu path e db per il salvataggio esterno.

todo: controllare se possiamo evitare di inserire path e db duplicati per lo stesso oggetti
      per cui si potrebbe costruire una chiave univca formata da:
         DISCO_DEST_s
         PATH_DEST_s
         DATABASE_DEST_s
         NOMEOGGETTO_s
      AL FINE DI EVITARE DUPLICATI DI SALVATAGGI ESTERNI
      provare la chiave composta....




CHATGPT

   SEGNALIBRI
      TI POSSO DARE ...
      
