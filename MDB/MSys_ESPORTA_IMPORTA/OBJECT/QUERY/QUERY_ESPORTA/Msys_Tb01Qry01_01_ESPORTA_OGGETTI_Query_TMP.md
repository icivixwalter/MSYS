Msys_Tb01Qry01_01_ESPORTA_OGGETTI_Query_TMP





## sql
      
      SELECT
             Msys_Tb01_ESPORTA_DB_EST.PATH_s,
             Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.NRO_OGGETTO_i,
             Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s,
             Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s,
             Msys_Tb01_ESPORTA_DB_EST.Scel_b
         FROM
             Msys_TB01_ESPORTA_DB_EST_TMP
             INNER JOIN (
                 Msys_Tb01_ESPORTA_DB_EST
                 INNER JOIN Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB ON (
                     Msys_Tb01_ESPORTA_DB_EST.COD_PROGETTO_s = Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.COD_PROGETTO_s
                 )
                 AND (
                     Msys_Tb01_ESPORTA_DB_EST.PATH_s = Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.PATH_s
                 )
             ) ON (
                 Msys_TB01_ESPORTA_DB_EST_TMP.PATH_s = Msys_Tb01_ESPORTA_DB_EST.PATH_s
             )
             AND (
                 Msys_TB01_ESPORTA_DB_EST_TMP.COD_PROGETTO_s = Msys_Tb01_ESPORTA_DB_EST.COD_PROGETTO_s
             )
         WHERE
             (
                 (
                     (
                         Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s
                     ) = "QUERY"
                 )
                 AND (
                     (
                         Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s
                     ) > ""
                 )
             )
         WITH
             OWNERACCESS
         OPTION;
