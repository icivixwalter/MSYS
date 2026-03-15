# QUERY SELECT DELL'ELENCO DEGLI OGGETTI DEL PROGETTO IN ESAME


## DENOMINAZIONE
La query che visualizza gli oggetti del database in questione è:
   PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti



## sql

   SELECT
         "PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti"         AS [TxtQuery],
         "3 CHIAVI DELLA TABELLA ELENCO OGGETTI --->"            AS [TxtLeChiavi],
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.COD_PROGETTO_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.TIPOGGETTO_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NOMEOGGETTO_s,

         "campi ---->"                                           AS [EX__________________________________],
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NRO_OGGETTO_i,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NOTA_OGGETTO_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NOTEex_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.EXPORT_b,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.CODICE_QUOTE_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NOTE_PROGETTO_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Expression_m,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Flag_b,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.LvExtra_i,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Name2_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.ObjectId_lng,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Order_b,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NOTE_ATTRIBUTI_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.PROPRIETA,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.VALOREPROPRIETA,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.COMANDO_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Scel_b,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Stam_b,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Canc_b,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Agg_b,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.BLOCC_b,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.IdOrd_lng,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.IDKeyArch_lng,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.Msg01_s,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.DATAINS_d,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.DATAAGG_d,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.TIMEOPER_d,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.ORA_AGG_d,
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.ID_lng,
          "Elenco oggetti del progetto" AS Tx01,
          "le 4 chiavi"                                           AS [TxtChiavi],
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.COD_PROGETTO_s    AS [KeyCOD_PROGETTO_s],
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.TIPOGGETTO_s      AS [KeyTIPOGGETTO_s],
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.COD_PROGETTO_s    AS [ORD_COD_PROGETTO],
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NRO_OGGETTO_i     AS [ORD_NRO_OGGETTO_i] 
      FROM
          PROGETTI_Msys_TB03_ELENCO_ObjProgetti;
