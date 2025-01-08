MSys_MACRO_Qry01_01_SELECT_TUTTI



  Note
        Query di estrazione di tutte le MACRO del PROGETTO MSYS .  La estrazione viene limitata
        solo agli oggetti tabella del PROGETTO MSYS con il filtro [COD_PROGETTO_s])="Msys"

    struttura   
     
        SELECT 
        MSys_MACRO.*, "TxtORDINAMENTO --->" AS txtORD, 
        [MSys_MACRO].[COD_PROGETTO_s] AS ORD_COD_PROGETTO_s, 
        "TxtCodici --->" AS txtCodici, 
        [MSys_MACRO].[NRO_OGGETTO_i] AS ORD, 
        [MSys_MACRO].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s
        
        FROM 
        MSys_MACRO
        
        WHERE 
        ((([MSys_MACRO].[COD_PROGETTO_s])="Msys"))
        
        ORDER BY 
        [MSys_MACRO].[COD_PROGETTO_s], [MSys_MACRO].[NRO_OGGETTO_i];
