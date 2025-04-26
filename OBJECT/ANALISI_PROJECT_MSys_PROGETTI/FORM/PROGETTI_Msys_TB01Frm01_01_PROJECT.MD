PROGETTI_Msys_TB01Frm01_01_PROJECT.MD
   Note
      form gestione progetti.


   CAMPO
      COD_PROGETTO_s_TXT
            Con il campo @codice@progetto si possono attivare questi eventi:
         
         EVENTO_DOPPIO_CLICK_(COD_PROGETTO_s_TXT_DblClick)

            con questo evento doppio click sul codice progetto viene salvato nella tabella
            tmp (PROGETTI_Msys_TB01_PROJECT_TMP) il codice del progetto che poi viene recuperato e visualizzato nella casella di testo filtro.
               codice---> 01_COD_PROGETTO_s_TXT_DblClick
               
         EVENTO_CHANGE
            con l'evento modifica del campo codice progetto aggiorno 
            anche i campi testo prelevando i dati dalla combinata con
            indice 1 e 2 e 4:
               PROGETTO_s_TXT 
               NOTE_PROGETTO_s_Txt
               GROUP_COD_PROG_s_TXT
               
                  codice---> 03_COD_PROGETTO_s_TXT.CHANGE

