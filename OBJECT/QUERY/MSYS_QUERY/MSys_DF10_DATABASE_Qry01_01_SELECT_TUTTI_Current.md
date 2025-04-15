MSys_DF10_DATABASE_Qry01_01_SELECT_TUTTI_Current
    
    Note
        Query che estra il progetto corrente con il filtro codice progetto
    schema
        SELECT MSys_DF10_DATABASE.*
            FROM 
            MSys_DF10_DATABASE 
            INNER JOIN 
            MSys_ELENCO_ObjProgetti_TMP 
            ON 
            MSys_DF10_DATABASE.COD_PROGETTO_s = MSys_ELENCO_ObjProgetti_TMP.COD_PROGETTO_s;

