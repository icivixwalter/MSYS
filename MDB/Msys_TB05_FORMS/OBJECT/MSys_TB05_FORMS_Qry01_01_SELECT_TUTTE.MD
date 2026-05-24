
# QUERY DI SELEZIONE DI TUTTE LE FORM SALVATE IN TABELLA

## NOTE
il file di gestione delle form salvate in tabella è gestito con la seguente
query: MSys_TB06_FORMS_Qry01_01_SELECT_TUTTE






## codic access
   codice ----> MSys_TB06_FORMS_Qry01_(query di estrazione delle form salvate i tabella di archivio)
   codice ----> MSys_TB06_FORMS_Qry0

## SQL
   SELECT
       [MSys_TB06_FORMS].[COD_PROGETTO_s]  AS [ORD_COD_PROGETTO_s],
       [MSys_TB06_FORMS].[NRO_OGGETTO_i]   AS [ORD],
       [MSys_TB06_FORMS].[TIPOGGETTO_s],
       [MSys_TB06_FORMS].[NOMEOGGETTO_s],
       [MSys_TB06_FORMS].[COD_ACCESS_S],
       [MSys_TB06_FORMS].[NOTA_OGGETTO_s],
       [MSys_TB06_FORMS].[NOTEex_s],
       [MSys_TB06_FORMS].[EXPORT_b],
       [MSys_TB06_FORMS].[COD_PROGETTO_s],
       [MSys_TB06_FORMS].[DISCO_s],
       [MSys_TB06_FORMS].[PATH_s],
       [MSys_TB06_FORMS].[DATABASE_s],
       [MSys_TB06_FORMS].[Attribute_by],
       [MSys_TB06_FORMS].[Expression_m],
       [MSys_TB06_FORMS].[Flag_b],
       [MSys_TB06_FORMS].[LvExtra_lng],
       [MSys_TB06_FORMS].[Name1_s],
       [MSys_TB06_FORMS].[Name2_s],
       [MSys_TB06_FORMS].[ObjectId_lng],
       [MSys_TB06_FORMS].[Order_bn],
       [MSys_TB06_FORMS].[NOTE_s],
       [MSys_TB06_FORMS].[TIPO_s],
       [MSys_TB06_FORMS].[PROPRIETA_s],
       [MSys_TB06_FORMS].[VALOREPROPRIETA_m],
       [MSys_TB06_FORMS].[COMANDO_s],
       [MSys_TB06_FORMS].[Scel_b],
       [MSys_TB06_FORMS].[Stam_b],
       [MSys_TB06_FORMS].[Canc_b],
       [MSys_TB06_FORMS].[Agg_b],
       [MSys_TB06_FORMS].[BLOCC_b],
       [MSys_TB06_FORMS].[IdOrd_lng],
       [MSys_TB06_FORMS].[IDKeyArch_lng],
       [MSys_TB06_FORMS].[Msg01_s],
       [MSys_TB06_FORMS].[DATAINS_d],
       [MSys_TB06_FORMS].[DATAAGG_d],
       [MSys_TB06_FORMS].[TIMEOPER_d],
       [MSys_TB06_FORMS].[ORA_AGG_d],
       [MSys_TB06_FORMS].[ID_FRM_lng]
      FROM
       MSys_TB06_FORMS
      ORDER BY
       [MSys_TB06_FORMS].[COD_PROGETTO_s],
       [MSys_TB06_FORMS].[NRO_OGGETTO_i];




## faq
   la query di estrazione di tutte le form salvate in tabella
   @query@estrai@forms
