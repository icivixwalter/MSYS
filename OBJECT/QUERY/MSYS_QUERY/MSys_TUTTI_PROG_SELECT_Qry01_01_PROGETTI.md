MSys_TUTTI_PROG_SELECT_Qry01_01_PROGETTI
    

    Note
        query di estrazione oggetti tabelle tmp
            @query@estrai@tutti.progetti
            @query@estrai@progetti.tutti

    struttura

            SELECT 
            MSys_TUTTI_PROG.*, 
            "->" AS TxtChiavi, 
            MSys_TUTTI_PROG.COD_PROGETTO_s AS [KeyCOD_PROGETTO_s] 
            FROM MSys_TUTTI_PROG
            ORDER BY MSys_TUTTI_PROG.COD_PROGETTO_s;


        
        
