CONFRONTA_Qry01_TB03_ELENCO_ObjProgetti.MD
   NOTA
      query chiamata confronta che estrae tutti i campi i comune con 
      la query CONFRONTA_Qry01_ESPORTA_IMPORTA

   SELECT 
      "PROGETTI" AS TXT01, 
      PROGETTI_Msys_TB03_ELENCO_ObjProgetti.COD_PROGETTO_s, 
      PROGETTI_Msys_TB03_ELENCO_ObjProgetti.PROGETTO_s, 
      PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NRO_OGGETTO_i, 
      PROGETTI_Msys_TB03_ELENCO_ObjProgetti.TIPOGGETTO_s, 
      PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NOMEOGGETTO_s, 
      PROGETTI_Msys_TB03_ELENCO_ObjProgetti.NOTA_OGGETTO_s, 
      PROGETTI_Msys_TB03_ELENCO_ObjProgetti.EXPORT_b

   FROM 
      PROGETTI_Msys_TB03_ELENCO_ObjProgetti
      WITH OWNERACCESS OPTION;
