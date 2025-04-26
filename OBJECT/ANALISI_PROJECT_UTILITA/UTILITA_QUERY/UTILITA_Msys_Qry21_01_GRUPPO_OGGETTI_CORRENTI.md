UTILITA_Msys_Qry21_01_GRUPPO_OGGETTI_CORRENTI

      NOTE
         elenco oggetti correnti


      SCHEMA

         SELECT 
            MSysObjects.Type, MSysObjects.Type AS kEYType, 
            MSysObjects.ParentId AS KEYParentId, 
            IIf([TYPE]=-32768,"FORM",
            IIf([TYPE]=-32766,"MACRO",
            IIf([TYPE]=-32764,"REPORT",
            IIf([TYPE]=-32761,"MODULI",
            IIf([TYPE]=-32757,"AccessLayout",
            IIf([TYPE]=1,"MSysObjects",
            IIf([TYPE]=2,"MSysDb",
            IIf([TYPE]=3,
            "Tutti gli oggetti Access",IIf([TYPE]=5,"QUERY + QUERY CANCELLATE",IIf([TYPE]=6,"TABELLE","ALTRO")))))))))) AS TIPO, 

            Count(MSysObjects.Type) AS NRO_c
            FROM MSysObjects
            GROUP BY MSysObjects.Type, MSysObjects.ParentId
            WITH OWNERACCESS OPTION;




