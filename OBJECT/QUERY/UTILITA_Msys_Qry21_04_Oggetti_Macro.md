UTILITA_Msys_Qry21_04_Oggetti_Macro.md 



   Nota
      estraggo solo le MACRO valide dalla tabella msys registrate nel db corrente


   SCHEMA

      SELECT 
         MSysObjects.Type AS kEYType, MSysObjects.ParentId AS KEYParentId, "--->" AS TXTCampi, MSysObjects.Type, MSysObjects.Connect, MSysObjects.Database, MSysObjects.Name
         FROM MSysObjects
         WHERE (((MSysObjects.Type)=-32766))
         WITH OWNERACCESS OPTION;




