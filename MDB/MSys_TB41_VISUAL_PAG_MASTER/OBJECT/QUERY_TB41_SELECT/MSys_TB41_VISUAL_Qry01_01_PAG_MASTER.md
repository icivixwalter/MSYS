# QUERY DI ESTRAZIONE DALLA @TABELLA TB41 VISUAL @MASTER

## ATTIVITA
La query 
   MSys_TB41_VISUAL_Qry01_01_PAG_MASTER.md
estrae i dati realtivi alla form da visualizzare dalla tabella master per il progetto
generale chiamato master. I dati vengono estratti dalla tabella MSys_TB41_VISUAL_PAG_MASTER


## SQL

   SELECT
       MSys_TB41_VISUAL_PAG_MASTER.*,
             "-->" AS TxtChiavi,
             [MSys_TB41_VISUAL_PAG_MASTER].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s,
             "-->" AS TxtORDi,
             [MSys_TB41_VISUAL_PAG_MASTER].[NRO_OGGETTO_i] AS ORD_NRO_OGGETTO_i
         FROM
             MSys_TB41_VISUAL_PAG_MASTER
         ORDER BY
             [MSys_TB41_VISUAL_PAG_MASTER].[NRO_OGGETTO_i];
      


