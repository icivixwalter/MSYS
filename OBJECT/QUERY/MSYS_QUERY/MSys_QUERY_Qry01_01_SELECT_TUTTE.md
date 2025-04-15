MSys_QUERY_Qry01_01_SELECT_TUTTE


  Note
        Query di estrazione di tutte le QUERY del PROGETTO MSYS .  La estrazione viene limitata
        solo agli oggetti tabella del PROGETTO MSYS con il filtro [COD_PROGETTO_s])="Msys"

    struttura   
     
        SELECT 
        MSys_QUERY.*, "TxtORDINAMENTO --->" AS txtORD, 
        [MSys_QUERY].[COD_PROGETTO_s] AS ORD_COD_PROGETTO_s, 
        "TxtCodici --->" AS txtCodici, 
        [MSys_QUERY].[NRO_OGGETTO_i] AS ORD, 
        [MSys_QUERY].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s
        
        FROM 
        MSys_QUERY
        
        WHERE 
        ((([MSys_QUERY].[COD_PROGETTO_s])="Msys"))
        
        ORDER BY 
        [MSys_QUERY].[COD_PROGETTO_s], [MSys_QUERY].[NRO_OGGETTO_i];
