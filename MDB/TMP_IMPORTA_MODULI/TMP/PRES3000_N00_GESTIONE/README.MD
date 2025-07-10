PRESS3000_N000_GESTIONE



MODIFICHE
   MACRO
      2025_07_10 = rinomino le macro
         COLLEGA_Mcr_PRES3000_DF01_CalendarioFestivita --->  PRES3000_Mcr02_DF01_CalendarioFestivita_Collega
         COLLEGA_Mcr_PRES3000_DF10_GIORNALIERE         --->  PRES3000_Mcr01_DF10_GIORNALIERE_Collega
         COLLEGA_Mcr_PRES3000_DF11_TIMBRATURE          --->  PRES3000_Mcr03_DF11_TIMBRATURE_Collega
         COLLEGA_Mcr_PRES3000_DF12_ORARI_TIMBRATURE    --->  PRES3000_Mcr04_DF12_ORARI_TIMBRATURE_Collega
         COLLEGA_Mcr_PRES3000_DF20_Help_Calendario     --->  PRES3000_Mcr05_DF20_Help_Calendario_Collega




IMPORTA OGGETTI
   Nota
      per importare gli oggetti viene utilizzata la fom msys UTILITA_Msys_Frm21_IMPORTA_OGGETTI (e la UTILITA_Msys_Frm21_frmProgressBar + 
      UTILITA_Msys_Frm21_UTILITA_ALTRI_PROGETTI) con la quale viene attivato il procedimento di importazione per gli oggetti 
      del PRES3000, le attivita sequenziali eseguite sono:

      LE ATTIVITA SEQUENZIALI DA ATTIVARE:
         
         @STEP@01_IMPORTA@OGGETTI_(reset la Cmb_01 NOME PROGETTO)          = Reset di default la Cmb_01 con il NOME DEL PROGETTO
         @STEP@01_IMPORTA@OGGETTI_(reset la Cmb_02)                        = Reset di default la Cmb_02
         @STEP@02_IMPORTA@OGGETTI_(Cmb_01 valori casella)                  = Evento della Cmb_01 imposto i valori della casella 
                     ... con il nome del progetto costruisci la funzione principale nella cmb_01 con
                         il menu del Pres3000
         @STEP@03_IMPORTA@OGGETTI_(Imposto la collection tables)           = Oggetti Tabella da importare
         @STEP@04_IMPORTA@OGGETTI_(Imposto la collection Macros)           = Oggetti Macro da importare
         @STEP@05_IMPORTA@OGGETTI_(Imposto la collection Modules)          = Oggetti Moduli da importare
         @STEP@06_IMPORTA@OGGETTI_(Imposto la collection Reports)          = Oggetti Reports da importare
         @STEP@07_IMPORTA@OGGETTI_(Abilito i Comandi)                      = Abilito i button Importa e esporta
                     ... con il codice IMPORTA_OGGETTI_PRES3000 abilito i button o i comandi di importazione
                         ed esportazione.

                     ... ATTENZIONE l'importazione o meglio il collegamento avviene in questa riga:
                           @14_@IMPORTA@TABELLE@PROG@IMPORT
                        controllare qui i vari errori.

         
            Nella combinata 02 definisco il nome della funzione:
               @IMPORTA@OGGETTI@PRES3000


            Gli oggetti da importare creati sono:
                  LE TABELLE

                      '//@COLLEZIONE@TABELLE@OGGETTI@PRES3000
                      '//@PRES3000@TABELLE@UTILITA_(LA @collection DELLE @TABELLE relative agli PRES3000)
                              Case "IMPORTA_OGGETTI_PRES3000"

                     con questo comando importo le tabelle ma devo cambiare la path per ogni tipo di tabella:
                           TABELLE DEFINIZIONI ORARI
                              PATH: c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB\PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB.mdb

                                 Tables_Collection.Add "PRES3000_DF00_{@==========================================@}_Tab"
                                 Tables_Collection.Add "PRES3000_DF01_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF01_CalendarioFestivita"
                                 Tables_Collection.Add "PRES3000_DF05_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF05_ORARI"
                                 Tables_Collection.Add "PRES3000_DF10_GIORNALIERE"
                                 Tables_Collection.Add "PRES3000_DF11_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF11_TIMBRATURE"
                                 Tables_Collection.Add "PRES3000_DF12_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF12_ORARI_TIMBRATURE"
                                 Tables_Collection.Add "PRES3000_DF19_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF19_GiorniCalendario"
                                 Tables_Collection.Add "PRES3000_DF20_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF20_Help_Calendario"

                           TABELLE CALENDARIO
                                 PATH: c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N01_CALENDARIO\MDB\PRES3000_N01_CALENDARIO.mdb

                                 Tables_Collection.Add "PRES3000_DF01_CalendarioFestivita"
                                 Tables_Collection.Add "PRES3000_DF19_GiorniCalendario"
                                 Tables_Collection.Add "PRES3000_DF20_Help_Calendario"
                                 Tables_Collection.Add "PRES3000_Tb01_Calendario"
                                 Tables_Collection.Add "PRES3000_Tb02_CalendarioPianoFerie"
               