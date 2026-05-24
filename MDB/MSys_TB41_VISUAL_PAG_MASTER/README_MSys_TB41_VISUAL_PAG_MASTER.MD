# PROGETTO TB41 VISUALIZZA PAGINE DEL PROGETTO MASTER

## ATTIVITA
con il progetto README_MSys_TB41_VISUAL_PAG_MASTER.MD 
gestisco TUTTE LE PAGINE DA VISUALIZZARE per i progetti 
gestiti



## CODICE OGGETTI ED OGGETTI COSTRUITI PER LA GESTIONE TB41
   CODICE---> MSys_TB41                = visualizzo tutti gli oggetti TB41
   CODICE---> MSys_TB41_VISUAL         = visualizzo solo le query TB41


TB41
   TABLE_TB41
      MSys_TB41_}-----------------------------@PAG_MASTER
      MSys_TB41_VISUAL_PAG_MASTER                        = Tabella che gestisce le form dei progetti
      MSys_TB41_VISUAL_PAG_MASTER_TMP                    = Tabella che gestisce le FORM DEL PROGETTO CORRENTE salvata in questa tmp

   QUERY_TB41_SELECT
       MSys_TB41_SELECT_Qry01_01_Archivio                = estraggo i progetti gestiti
   QUERY_TB41_DELETE
         MSys_TB41_DELETE_Qry01_01_ArchivioTMP           = CANCELLO I DATI TMP
   QUERY_TB41_INSERT
      MSys_TB41_INSERT_Qry01_01_Archivio*
            
   QUERY_TB41_UPDATE
         ......
         MSys_TB41_UPDATE_Qry01_01_NRO_PROG_i            = aggorno il nro progetto nella tab TB41

                

   
   FORM
      MSys_TB41_Frm01_S01_Archivio                       = FORM dei progetti gestiti
