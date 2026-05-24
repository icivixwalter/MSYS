# @README PROGETTO @TB4O@PROGETTO@MASTER GESTIONE DEI PROGETTI MASTER

## ATTIVITA
con il progetto Msys_TB40_PROG_MASTER.MD costruisco il PROGETTO MASTER di
gestione delle form collegate al progetto. Per ogni progetto per ora sono
collegate le form da visualizzare con la tabella sottostante collegata.


Per ora salvo il progetti con il codice gia
definito nella tabella progetti in questi campi:  
    NRO_PROG_i                          = numero di progetto gia definito
    COD_PROGETTO_s                      = codice progettto gestito
    DENOM_PROGETTO_s                    = denominazione progetto.

Per la gestione dei progetti sono stati costruiti gli oggetti 
TABLE, QUERIES E FORMS

## PROPRIETA
La tabella master è collegata in una sottotabella detail con la quale individuare
per ogni singolo progetto quali oggetti form sono necessari per il funzionamento i
base a questa relazione :
        MSys_TAB40_PROG_MASTER  (master)
                |
                |key Master/detail = COD_PROGETTO_s
                |
        MSys_TB41_VISUAL_PAG_MASTER (detail)

In questo modo per il progetto A sono visibile le form collegate:
        MSys_Progetto_01
            |
            |------ FORM A)
            |------ FORM B)
            |------ FORM C)
            |------ (....)
            |




## CODICI ACCESS PROGETTO TB40
      codice ----> TB40                               = tutti gli oggetti del progetto TB40
      codice ----> TAB40_Qry01                        = ESTRAGGO gli oggetti QUERY del progetto

## OGGETTI COSTRUITI
   TABLES
      MSys_TAB40_}----------------------------------------@FORM_MASTER
      MSys_TAB40_FORM_MASTER                                = tabella dove vengono salvate i progetti master
      MSys_TAB40_FORM_MASTER_TMP
          
      MSys_TB40_SELECT_Qry01_01_MASTER                      = Query di visualizazione delle pagine
      MSys_TB40_INSERT_Qry01_01_Master_toTmp                = Query inserisci su tmp
      

         MSyTB40_PROG_MASTER_TMP_SELECT_Qry01_01_ArchivioTmp   = estraggo i progetti in tabellA TMP
         MSys_TB40_PROG_MASTER_SELECT_Qry01_01_Archivio        = estraggo i progetti master
            estrazione dalla tabella master 
         MSys_TB40_PROG_MASTER_INSERT_Qry01_01_toTmp           = inserisco in tmp i dati da archivio
            inserisce in tmp
         MSys_TB40_PROG_MASTER_DELETE_TMP_Qry01_01_Svuota      = svuoto tabella tmp
         MSys_TAB40_PROG_Qry01_01_MASTER_UPDATE_NRO_PROG_i     = Aggiorno numero di progetto
