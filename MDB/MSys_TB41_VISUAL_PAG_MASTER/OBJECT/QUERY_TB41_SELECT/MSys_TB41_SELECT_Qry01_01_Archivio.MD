# QUERY DI ESTRAZIONE DALLA @TABELLA@MASTER

## ATTIVITA
La query 
      MSys_TAB402_VISUAL_PAG_MASTER_Qry01_01_Archivio
estrae i dati dalla tabella master



## SQL
   SELECT
          MSys_TAB402_VISUAL_PAG_MASTER.*,
          "-->" AS TxtChiavi,
          [MSys_TAB402_VISUAL_PAG_MASTER].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s,
          "-->" AS TxtORDi,
          [MSys_TAB402_VISUAL_PAG_MASTER].[NRO_OGGETTO_i] AS ORD_NRO_OGGETTO_i
      FROM
          MSys_TAB402_VISUAL_PAG_MASTER
      ORDER BY
          [MSys_TAB402_VISUAL_PAG_MASTER].[NRO_OGGETTO_i];
