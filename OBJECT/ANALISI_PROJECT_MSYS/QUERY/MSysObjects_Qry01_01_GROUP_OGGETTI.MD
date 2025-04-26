


SCHEMA

    SELECT MSysObjects.Type, 
           MSysObjects.ParentId, 
           "--->" AS TXTKey, 
           MSysObjects.Type AS kEYType, 
           MSysObjects.ParentId AS KEYParentId, 
           IIf([TYPE]=-32768,"FORM",
               IIf([TYPE]=-32766,"MACRO",
                   IIf([TYPE]=-32764,"REPORT",
                       IIf([TYPE]=-32761,"MODULI",
                           IIf([TYPE]=-32757,"AccessLayout",
                               IIf([TYPE]=1,"MSysObjects",
                                   IIf([TYPE]=2,"MSysDb",
                                       IIf([TYPE]=3,"Tutti gli oggetti Access",
                                          IIf([TYPE]=5,"QUERY + QUERY CANCELLATE",
                                          IIf([TYPE]=6,"TABELLE","ALTRO")))))))))) AS TIPO
                FROM MSysObjects
                GROUP BY MSysObjects.Type, MSysObjects.ParentId
                WITH OWNERACCESS OPTION;
