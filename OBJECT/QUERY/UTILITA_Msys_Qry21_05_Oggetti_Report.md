UTILITA_Msys_Qry21_05_Oggetti_Report.md 



   Nota
      estraggo solo le REPORT valide dalla tabella msys registrate nel db corrente


   SCHEMA

      SELECT 
         MSysObjects.Type AS kEYType, MSysObjects.ParentId AS KEYParentId, "--->" AS TXTCampi, MSysObjects.Type, MSysObjects.Connect, MSysObjects.Database, MSysObjects.Name
         FROM MSysObjects
         WHERE (((MSysObjects.Type)=-32764))
         WITH OWNERACCESS OPTION;




