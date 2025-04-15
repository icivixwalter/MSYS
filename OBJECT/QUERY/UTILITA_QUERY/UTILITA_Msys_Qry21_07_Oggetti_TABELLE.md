UTILITA_Msys_Qry21_07_Oggetti_TABELLE.md 



   Nota
      estraggo solo le TABELLE valide dalla tabella msys registrate nel db corrente
      UTILITA_Msys_Qry21_00_{@======================================@}


   SCHEMA

      SELECT 
         MSysObjects.Type AS kEYType, MSysObjects.ParentId AS KEYParentId, "--->" AS TXTCampi, MSysObjects.Type, MSysObjects.Connect, MSysObjects.Database, MSysObjects.Name
         FROM MSysObjects
         WHERE (((MSysObjects.Type)=6))
         WITH OWNERACCESS OPTION;




