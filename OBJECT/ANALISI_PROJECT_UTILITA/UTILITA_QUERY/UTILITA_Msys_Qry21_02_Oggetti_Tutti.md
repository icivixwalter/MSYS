UTILITA_Msys_Qry21_02_Oggetti_Tutti
   NOTE 
      Query utilita che serve ad estrarre gli oggetti validi nel db corrente
   
   SCHEMA
      SELECT 
         MSysObjects.Type AS kEYType, 
         MSysObjects.ParentId AS KEYParentId, 
         "--->" AS TXTCampi, 
         MSysObjects.Type, 
         MSysObjects.Connect, 
         MSysObjects.Database, MSysObjects.Name
         FROM MSysObjects
         WITH OWNERACCESS OPTION;
