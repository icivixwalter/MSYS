# QUERY INSERT DA MSYOBJECT IN TB99

## NOTE
query inserimento dati
Msys_TB99_LINK_Qry99_70_INSERT_ArchivioTabelleLINK

## sql
INSERT INTO
       Msys_TB99_LINK (
           TYPE,
           [Connect],
           [Database],
           Name
       )
   SELECT
       MSysObjects_Qry01_06_SELECT_TABELLE_LINK.Type,
       MSysObjects_Qry01_06_SELECT_TABELLE_LINK.Connect,
       MSysObjects_Qry01_06_SELECT_TABELLE_LINK.Database,
       MSysObjects_Qry01_06_SELECT_TABELLE_LINK.Name
   FROM
       MSysObjects_Qry01_06_SELECT_TABELLE_LINK
       LEFT JOIN Msys_TB99_LINK ON MSysObjects_Qry01_06_SELECT_TABELLE_LINK.Name = Msys_TB99_LINK.Name
   WHERE
       (((Msys_TB99_LINK.Name) IS NULL));
       
