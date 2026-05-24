# QUERY DI ACCODAMENTI DEI PROGETTI MASTER SU TMP

## ATTIVITA
La query 
   MSys_TB40_PROG_MASTER_INSERT_Qry01_01_toTmp
ACCODA i dati dalla tabella master alla tabella master TMP



## SQL
   
  INSERT INTO
    MSys_TAB40_PROG_MASTER_TMP (
        NRO_PROG_i,
        COD_PROGETTO_s,
        COD_FORMS_s,
        DENOM_PROGETTO_s,
        Scel_b,
        NOTE_PROGETTO_s,
        DATAINS_d,
        DATAAGG_d,
        TIMEOPER_d
    )
   SELECT
       MSys_TAB40_PROG_MASTER.NRO_PROG_i,
       MSys_TAB40_PROG_MASTER.COD_PROGETTO_s,
       MSys_TAB40_PROG_MASTER.COD_FORMS_s,
       MSys_TAB40_PROG_MASTER.DENOM_PROGETTO_s,
       MSys_TAB40_PROG_MASTER.Scel_b,
       MSys_TAB40_PROG_MASTER.NOTE_PROGETTO_s,
       MSys_TAB40_PROG_MASTER.DATAINS_d,
       MSys_TAB40_PROG_MASTER.DATAAGG_d,
       MSys_TAB40_PROG_MASTER.TIMEOPER_d
   FROM
       MSys_TAB40_PROG_MASTER,
       MSys_TAB40_PROG_MASTER_TMP
   WHERE
       (
           ((MSys_TAB40_PROG_MASTER.Scel_b) = True)
           AND (
               (MSys_TAB40_PROG_MASTER_TMP.NRO_PROG_i) IS NULL
           )
           AND (
               (MSys_TAB40_PROG_MASTER_TMP.COD_PROGETTO_s) IS NULL
           )
       );
   
   
   
   
   
