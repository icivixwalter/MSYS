


# QUERY DI INSERIMENTO DALLA TABELLA DI ARCHIVIO A QUELLA TMP

## ATTIVITA
La query 
   MSys_TB41_Qry01_01_VISUAL_PAG_MASTER_Insert_TMP.md
Inserisco i dati dalla tabella di archivio alla tmp:
      MSys_TB41_VISUAL_PAG_MASTER ---> MSys_TB41_VISUAL_PAG_MASTER_TMP

## SQL
*** todo: non funziona
   INSERT INTO
          MSys_TB41_VISUAL_PAG_MASTER_TMP (
              NRO_PROG_i,
              COD_FORMS_s,
              COD_PROGETTO_s,
              NRO_OGGETTO_i,
              TIPOGGETTO_s,
              NOMEOGGETTO_s,
              NOMEpagina_s,
              VISIBLE_b,
              NOTA_OGGETTO_s,
              NOTEex_s,
              DATAINS_d,
              DATAAGG_d,
              TIMEOPER_d,
              ORA_AGG_d
          )
      SELECT
          [MSys_TB41_VISUAL_PAG_MASTER].[NRO_PROG_i],
          [MSys_TB41_VISUAL_PAG_MASTER].[COD_FORMS_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[COD_PROGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[NRO_OGGETTO_i],
          [MSys_TB41_VISUAL_PAG_MASTER].[TIPOGGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[NOMEOGGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[NOMEpagina_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[VISIBLE_b],
          [MSys_TB41_VISUAL_PAG_MASTER].[NOTA_OGGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[NOTEex_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[DATAINS_d],
          [MSys_TB41_VISUAL_PAG_MASTER].[DATAAGG_d],
          [MSys_TB41_VISUAL_PAG_MASTER].[TIMEOPER_d],
          [MSys_TB41_VISUAL_PAG_MASTER].[ORA_AGG_d]
      FROM
          (
              MSys_TB41_VISUAL_PAG_MASTER
              INNER JOIN MSys_TB41_VISUAL_PAG_MASTER ON [MSys_TB41_VISUAL_PAG_MASTER].[COD_PROGETTO_s] = [MSys_TB41_VISUAL_PAG_MASTER].[COD_PROGETTO_s]
          )
          LEFT JOIN MSys_TB41_VISUAL_PAG_MASTER_TMP ON (
              [MSys_TB41_VISUAL_PAG_MASTER].[TIPOGGETTO_s] = [MSys_TB41_VISUAL_PAG_MASTER_TMP].[TIPOGGETTO_s]
          )
          AND (
              [MSys_TB41_VISUAL_PAG_MASTER].[COD_PROGETTO_s] = [MSys_TB41_VISUAL_PAG_MASTER_TMP].[COD_PROGETTO_s]
          )
      GROUP BY
          [MSys_TB41_VISUAL_PAG_MASTER].[NRO_PROG_i],
          [MSys_TB41_VISUAL_PAG_MASTER].[COD_FORMS_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[COD_PROGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[NRO_OGGETTO_i],
          [MSys_TB41_VISUAL_PAG_MASTER].[TIPOGGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[NOMEOGGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[NOMEpagina_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[VISIBLE_b],
          [MSys_TB41_VISUAL_PAG_MASTER].[NOTA_OGGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[NOTEex_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[DATAINS_d],
          [MSys_TB41_VISUAL_PAG_MASTER].[DATAAGG_d],
          [MSys_TB41_VISUAL_PAG_MASTER].[TIMEOPER_d],
          [MSys_TB41_VISUAL_PAG_MASTER].[ORA_AGG_d],
          [MSys_TB41_VISUAL_PAG_MASTER_TMP].[COD_PROGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER_TMP].[TIPOGGETTO_s],
          [MSys_TB41_VISUAL_PAG_MASTER].[Scel_b]
      HAVING
          (
              (
                  (MSys_TB41_VISUAL_PAG_MASTER_TMP.COD_PROGETTO_s) IS NULL
              )
              AND (
                  (MSys_TB41_VISUAL_PAG_MASTER_TMP.TIPOGGETTO_s) IS NULL
              )
              AND ((MSys_TB41_VISUAL_PAG_MASTER.Scel_b) = True)
          )
      ORDER BY
          [MSys_TB41_VISUAL_PAG_MASTER].[NRO_OGGETTO_i]
      WITH
          OWNERACCESS
      OPTION;
