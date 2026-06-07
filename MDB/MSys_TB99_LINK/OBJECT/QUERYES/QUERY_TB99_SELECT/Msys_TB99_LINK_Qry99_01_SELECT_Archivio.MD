# QUERY TB99 LINK ARCHIVIO

## NOTE
con il file 
Msys_TB99_LINK_Qry99_01_SELECT_Archivio
gestisco l'interrogazione sulla tabella tb99 di archivio dei collegamenti



## SQL
   SELECT
          Msys_TB99_LINK.COD_PROGETTO_c,
          Msys_TB99_LINK.COD_SUB_PROGETTO_c,
          Msys_TB99_LINK.Type,
          Msys_TB99_LINK.Connect,
          Msys_TB99_LINK.Database,
          Msys_TB99_LINK.Name,
          Msys_TB99_LINK.TIPO_OGGETTO_c,
          Msys_TB99_LINK.TIPO_c,
          Msys_TB99_LINK.Verifica_s
      FROM
          Msys_TB99_LINK;
