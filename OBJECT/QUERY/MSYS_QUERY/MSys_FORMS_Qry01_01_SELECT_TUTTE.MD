MSys_FORMS_Qry01_01_SELECT_TUTTE


  Note
        Query di estrazione di tutte le FORM del PROGETTO MSYS .  La estrazione viene limitata
        solo agli oggetti tabella del PROGETTO MSYS con il filtro [COD_PROGETTO_s])="Msys"

    struttura   
     
        SELECT 
        MSys_FORMS.*, "TxtORDINAMENTO --->" AS txtORD, 
        [MSys_FORMS].[COD_PROGETTO_s] AS ORD_COD_PROGETTO_s, 
        "TxtCodici --->" AS txtCodici, 
        [MSys_FORMS].[NRO_OGGETTO_i] AS ORD, 
        [MSys_FORMS].[COD_PROGETTO_s] AS KeyCOD_PROGETTO_s
        
        FROM 
        MSys_FORMS
        
        WHERE 
        ((([MSys_FORMS].[COD_PROGETTO_s])="Msys"))
        
        ORDER BY 
        [MSys_FORMS].[COD_PROGETTO_s], [MSys_FORMS].[NRO_OGGETTO_i];
