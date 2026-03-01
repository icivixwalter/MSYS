Msys_Tb01Qry01_01_ESPORTA_OGGETTI_Moduli









## sql
      
      SELECT
                Msys_ESPORTA_IMPORTA_DB_DEST.PATH_s,
                Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.NRO_OGGETTO_i,
                Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s,
                Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.NOMEOGGETTO_s,
                Msys_ESPORTA_IMPORTA_DB_DEST.Scel_b
            FROM
                Msys_ESPORTA_IMPORTA_DB_DEST
                INNER JOIN Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB ON (
                    Msys_ESPORTA_IMPORTA_DB_DEST.COD_PROGETTO_s = Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.COD_PROGETTO_s
                )
                AND (
                    Msys_ESPORTA_IMPORTA_DB_DEST.PATH_s = Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.PATH_s
                )
            WHERE
                (
                    (
                        (
                            Msys_Tb01_ESPORTA_ELENCO_OGGETTI_DB.TIPOGGETTO_s
                        ) = "MODULES"
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
