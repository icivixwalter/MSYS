UTILITA_Msys_Qry21_08_Oggetti_QUERY.md 




   Nota
      estraggo solo le QUERY valide dalla tabella msys registrate nel db corrente
      Ho aggiunto il campo LvExtra per estrarre se > null le QUERY non cancellate.

   SCHEMA

      SELECT 
         MSysObjects.Type AS kEYType, MSysObjects.ParentId AS KEYParentId, "--->" AS TXTCampi, MSysObjects.Type, MSysObjects.Connect, MSysObjects.Database, MSysObjects.Name
         FROM MSysObjects
         WHERE (((MSysObjects.Type)=5) AND ((MSysObjects.LvExtra)>""))
         WITH OWNERACCESS OPTION;




