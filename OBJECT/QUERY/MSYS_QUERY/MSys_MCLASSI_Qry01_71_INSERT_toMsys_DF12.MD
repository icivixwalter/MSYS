MSys_MCLASSI_Qry01_70_}----------------------------------@INSERT
MSys_MCLASSI_Qry01_71_INSERT_toMsys_DF12


                Note
                Query di inserimento OGGETTI MODULI nella tabella ELENCO OGGETTI DI PROGETTO utili
                per l'esportazione in un altro db esterno.
                        faq
                            @ACCODA@OGGETTI.MODULI
                            @ACCODA@MSYS@MODULI
                            @INSERT@MODULI@MSYS

        Struttura
        
                INSERT 
                INTO 
                MSys_ELENCO_ObjProgetti 

                ( Scel_b, Name1_s, IdOrd_lng, 
                NRO_OGGETTO_i, TIPOGGETTO_s, 
                COD_PROGETTO_s, DATAINS_d, TIMEOPER_d )

                SELECT 
                
                MSys_MCLASSI.Scel_b, 
                MSys_MCLASSI.Name1_s, 
                MSys_MCLASSI.ID_MCLASSI_lng, 
                MSys_MCLASSI.NRO_OGGETTO_i, 
                "MODULI" AS TIPOGGETTO_s, 
                MSys_MCLASSI.COD_PROGETTO_s,
                MSys_MCLASSI.DATAINS_d, MSys_MCLASSI.TIMEOPER_d
          
                FROM 
                MSys_MCLASSI LEFT JOIN MSys_ELENCO_ObjProgetti 
                ON (MSys_MCLASSI.COD_PROGETTO_s = MSys_ELENCO_ObjProgetti.COD_PROGETTO_s) 
                AND (MSys_MCLASSI.Name1_s = MSys_ELENCO_ObjProgetti.Name1_s)
               
                WHERE 
                (((MSys_MCLASSI.Scel_b)=True) 
                AND ((MSys_ELENCO_ObjProgetti.Name1_s) Is Null) 
                AND ((MSys_ELENCO_ObjProgetti.COD_PROGETTO_s) Is Null))
                WITH OWNERACCESS OPTION;
