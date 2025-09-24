# README_MINI_PROGETTO_Msys_TB01_PROGETTI

### GE_Msys_TB01_PROGETTI.mdb
   NOTE
      progetto DI GESTIONE DI TUTTI PROGETTI che viene eseguito mediante il PROGETTO
      ACCCESS: 
         GE_Msys_TB01_PROGETTI.mdb
      In esso vengono gestite :
         a) la definizione del codice progetto con l'utilizzo deglio oggetti del mini progetto
            DF11.
               FORM
                  Msys_DF11_Frm01_M01_GE_CODICE_PROGETTI_TUTTI
                     si trova in questa path:
   
                        start "apri PROGETTI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB01_PROGETTI\MDB\GE_Msys_TB01_PROGETTI.mdb"

         B) i PROGETTI 
            ossia tutti gli oggetti per gestire il singolo MINI PROGETTO in 
               start "apri gestione oggetti dei MINI PROGETTI" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB01_PROGETTI\MDB\GE_Msys_TB01_PROGETTI.mdb"
   CORREZIONE FORM --> PROGETTI_Msys_TB02Frm01_01_TIPO_OGGETTI
               NEL DB : GE_Msys_TB01_PROGETTI.mdb
                           |
                           | --->FORM -> PROGETTI_Msys_TB02Frm01_01_TIPO_OGGETTI
                                                |
                                                |__Non funziona nel campo ---> TIPO_OGGETTO = risolto
   CORREZIONE FORM --> PROGETTI_Msys_TB03Frm01_01_ELENCO_ObjProgetti
               NEL DB : GE_Msys_TB01_PROGETTI.mdb
                              |
                              |____________FORM
                                           |
                                           |__PROGETTI_Msys_TB03Frm01_01_ELENCO_ObjProgetti
                                                      |  
                                                      |__COD_PROGETTO = casella di controllo non funziona.
                                                                        risolto
   CORREZIONE FORM --> Msys_DF11_Frm01_S01_COD_PROGETTI
               NEL DB : GE_Msys_TB01_PROGETTI.mdb
                              |
                              |____________FORM
                                           |
                                           |__Msys_DF11_Frm01_S01_COD_PROGETTI
                                              |    
                                              |-----PATH_DEST_s = casella di controllo ammessa a tutti gli elenchi



   CORREZIONE FORM --> Msys_DF01_Frm01_M01_GE_COMANDI
               NEL DB : PATH apri con start "APRI PROGETTO COMANDI" CALL "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\GE_MSys_DF01.mdb"
                              |
                              |____________FORM
                                           |
                                           |__Msys_DF01_Frm01_M01_GE_COMANDI
                                              |    
                                              |-----COD_PROGETTO_s_txt     = casella di controllo impostata query
                                              |-----PROGETTO_s_TXT         = casella di controllo impostata query
                                              |-----COMANDO_s_TXT          = casella di controllo impostata query
                                              |-----per tutte le caselle combinate: SELECT MSys_DF01_COMANDI.COMANDO_s FROM MSys_DF01_COMANDI; 





CHATGPT 
    TODO: da fare...

