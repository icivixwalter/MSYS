# QUERY DI ESTRAZIONE DALLA @TABELLA@PROGETTO@MASTER

## ATTIVITA
La query 
  MSys_TB40_SELECT_Qry01_}------------------------------------@Qry
   MSys_TB40_SELECT_Qry01_01_MASTER
estrae i dati dalla tabella relativa al progetto master
MSys_TB40_SELECT_Qry01_01_MASTER



## SQL
   
  SELECT
          MSys_TAB40_PROG_MASTER_TMP.*,
          "->" AS TxtChiavi,
          MSys_TAB40_PROG_MASTER_TMP.COD_PROGETTO_s AS [KeyCOD_PROGETTO_s] 
      FROM
          MSys_TAB40_PROG_MASTER_TMP;



## PROPRIETA




## FAQ

TAB40_Qry01_(il codice delle query relative al progetto master)                           
