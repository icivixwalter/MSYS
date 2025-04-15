MSys_MODULI_Qry01_01_SELECT_TUTTI


 Note
        Query di estrazione di I REPORT del PROGETTO MSYS .  La estrazione viene limitata
        solo agli oggetti tabella del PROGETTO MSYS con il filtro [COD_PROGETTO_s])="Msys"

    struttura   
     
        SELECT 
        MSys_MODULI.*, "TxtORDINAMENTO --->" AS txtORD, 
        [MSys_MODULI].[COD_PROGETTO_s] AS ORD_COD_PROGETTO_s, 
        "TxtCodici --->" AS txtCodici, 
        [MSys_MODULI].[NRO_OGGETTO_i] AS ORD, 
        [MSys_MODULI].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s
        
        FROM 
        MSys_MODULI
        
        WHERE 
        ((([MSys_MODULI].[COD_PROGETTO_s])="Msys"))
        
        ORDER BY 
        [MSys_MODULI].[COD_PROGETTO_s], [MSys_MODULI].[NRO_OGGETTO_i];
