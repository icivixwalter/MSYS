  # QUERY DI ESTRAZIONE DALLA @TABELLA@PROGETTO@MASTER@TMP

## ATTIVITA
La query 
   MSys_TB40_PROG_MASTER_DELETE_TMP_Qry01_01_Svuota
estrae i dati dalla tabella relativa al progetto master



## SQL
   
  DELETE
          MSys_TAB40_PROG_MASTER_TMP.*,
          "->" AS TxtChiavi,
          MSys_TAB40_PROG_MASTER_TMP.COD_PROGETTO_s AS [KeyCOD_PROGETTO_s] 
      FROM
          MSys_TAB40_PROG_MASTER_TMP;


