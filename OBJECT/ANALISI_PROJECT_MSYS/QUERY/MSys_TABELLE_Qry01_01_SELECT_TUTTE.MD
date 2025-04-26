MSys_TABELLE_Qry01_01_SELECT_TUTTE



   Note
        Query di estrazione di tutte le TABELLE del PROGETTO MSYS .  La estrazione viene limitata
        solo agli oggetti tabella del PROGETTO MSYS con il filtro [COD_PROGETTO_s])="Msys"

    struttura   
     
        SELECT 
        MSys_TABELLE.*, "TxtORDINAMENTO --->" AS txtORD, 
        [MSys_TABELLE].[COD_PROGETTO_s] AS ORD_COD_PROGETTO_s, 
        "TxtCodici --->" AS txtCodici, 
        [MSys_TABELLE].[NRO_OGGETTO_i] AS ORD, 
        [MSys_TABELLE].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s
        
        FROM 
        MSys_TABELLE
        
        WHERE 
        ((([MSys_TABELLE].[COD_PROGETTO_s])="Msys"))
        
        ORDER BY 
        [MSys_TABELLE].[COD_PROGETTO_s], [MSys_TABELLE].[NRO_OGGETTO_i];



  