# QUERY DI AGGIORNAMENTO NRO PROGETTO NELLA TABELLA TB41

## ATTIVITA
La query 
   MSys_TB40_PROG_MASTER_UPDATE_Qry01_01_NRO_PROG_i

Aggiorna il numero di progetto nella tabella MSys_TB41_VISUAL_PAG_MASTER di
visualizzazione delle pagine del progetto master.



## SQL
      
   UPDATE 
            MSys_TAB40_PROG_MASTER INNER JOIN 
            MSys_TB41_VISUAL_PAG_MASTER ON 
            MSys_TAB40_PROG_MASTER.COD_PROGETTO_s = MSys_TB41_VISUAL_PAG_MASTER.COD_PROGETTO_s 
            SET 
            MSys_TB41_VISUAL_PAG_MASTER.NRO_PROG_i = [MSys_TAB40_PROG_MASTER].[NRO_PROG_i]
         WITH 
            OWNERACCESS OPTION;



## FAQ
come aggiornare il numero di progetto?
@aggiorna@nro@progetto_(AGGIORNO il nro del progetto nella tabella TB41)
