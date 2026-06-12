# query DI BACKUP

## NOTE
con il file
Msys_TB99_LINK_Qry99_01_SELECT_BackUp
gestisco la query per il backup dei dati in un database esterno.

 SELECT
          Msys_TB99_LINK.Name,
          Msys_TB99_LINK.TIPO_OGGETTO_c,
          Msys_TB99_LINK.Verifica_s,
          Msys_TB99_LINK.Scel_b
      FROM
          Msys_TB99_LINK
      WHERE
          (
              (
                  (Msys_TB99_LINK.Name) NOT LIKE '*[_]{@*'
                  AND (Msys_TB99_LINK.Name) NOT LIKE '*[_]}*'
                  AND (Msys_TB99_LINK.Name) IS NOT NULL
                  AND (Msys_TB99_LINK.Name) <> ''
              )
          );
