MSys_REPORT_Qry01_01_SELECT_TUTTE


 Note
        Query di estrazione di I REPORT del PROGETTO MSYS .  La estrazione viene limitata
        solo agli oggetti tabella del PROGETTO MSYS con il filtro [COD_PROGETTO_s])="Msys"

    struttura   
     
        SELECT 
        MSys_REPORT.*, "TxtORDINAMENTO --->" AS txtORD, 
        [MSys_REPORT].[COD_PROGETTO_s] AS ORD_COD_PROGETTO_s, 
        "TxtCodici --->" AS txtCodici, 
        [MSys_REPORT].[NRO_OGGETTO_i] AS ORD, 
        [MSys_REPORT].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s
        
        FROM 
        MSys_REPORT
        
        WHERE 
        ((([MSys_REPORT].[COD_PROGETTO_s])="Msys"))
        
        ORDER BY 
        [MSys_REPORT].[COD_PROGETTO_s], [MSys_REPORT].[NRO_OGGETTO_i];
