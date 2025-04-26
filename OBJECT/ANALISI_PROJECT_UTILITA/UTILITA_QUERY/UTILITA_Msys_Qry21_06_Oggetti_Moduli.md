UTILITA_Msys_Qry21_06_Oggetti_Moduli.md 



   Nota
      estraggo solo le MODULI valide dalla tabella msys registrate nel db corrente


   SCHEMA

      SELECT 
         MSysObjects.Type AS kEYType, MSysObjects.ParentId AS KEYParentId, "--->" AS TXTCampi, MSysObjects.Type, MSysObjects.Connect, MSysObjects.Database, MSysObjects.Name
         FROM MSysObjects
         WHERE (((MSysObjects.Type)=-32761))
         WITH OWNERACCESS OPTION;




