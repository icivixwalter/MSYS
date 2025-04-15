MSys_TABELLE_Qry01_01_SELECT_CurrentProject

    Note
        Questa query estrae le tabelle del progetto corrente 
            AGGIORNAMENTO_CODICE---->2022.09.29@TABELLE@CURRENTPROJECT
                @TABELLE@PROGETTO.CORRENTE
                @modifca.2022.10.06 = ho modificato la query  in modo che
                faccia l'estrazione non dalla tabella TMP ma dai progetti tmp
                MSys_ELENCO_ObjProgetti_TMP
    
    struttura

        SELECT MSys_TABELLE.*, 
               MSys_TABELLE.COD_PROGETTO_s AS ORD_COD_PROGETTO_s, 
               MSys_TABELLE.NRO_OGGETTO_i AS ORD
               FROM MSys_TABELLE 
               INNER JOIN 
               MSys_ELENCO_ObjProgetti_TMP 
               ON MSys_TABELLE.COD_PROGETTO_s = MSys_ELENCO_ObjProgetti_TMP.COD_PROGETTO_s
               ORDER BY MSys_TABELLE.COD_PROGETTO_s, 
               MSys_TABELLE.NRO_OGGETTO_i;
