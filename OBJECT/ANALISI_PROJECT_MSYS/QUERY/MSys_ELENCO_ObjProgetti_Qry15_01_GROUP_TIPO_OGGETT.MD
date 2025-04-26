MSys_ELENCO_ObjProgetti_Qry15_01_GROUP_TIPO_OGGETTI


   Note
        Query di estrazione VISUALIZZAZIONE DEL GRUPPO DI PROGETTI 

    struttura   
     
            SELECT 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI].[COD_PROGETTO_s], 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI].[TIPOGGETTO_s], 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI].[PROGETTO_s], 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI].[NOTA_OGGETTO_s],
                "chiavi create"                         AS [TXkey], 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI].[COD_PROGETTO_s] 
                                                        AS [KeyCOD_PROGETTO_s]

            FROM 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI]

            GROUP BY 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI.COD_PROGETTO_s], 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI.TIPOGGETTO_s], 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI.PROGETTO_s], 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI.NOTA_OGGETTO_s],
                "chiavi create", 
                [MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI.COD_PROGETTO_s];





    FAQ

        @VISUALIZZA.I.PROGETTI
        @QUERY.PROGGETTI
        @GRUPPO@PROGETTI
