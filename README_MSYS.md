# SCHEMA CARTELLE
    Note
         schema delle cartelle e dei file sul percorso del progetto:
               c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\
    MSYS/
      │
      ├─GIT_BAT/
      ├─MDB/
      ├─OBJECT/
      ├─XLS/
      ├─.gitignore
      ├─A.TXT
      ├─AA________________________________________________________________________AA.txt
      ├─ATTIVA_SALVATAGGI_COLLETTIVI_MSYS_OGGETTI+MENU+IMPEGNI.BAT
      ├─CANCELLA_{MDB+OBJECT+PROJECT}.BAT
      ├─git_OPEN_OVERLEAF.BAT
      ├─git_OpenGitHub.bat
      ├─git_PULL.BAT
      ├─git_STATUS.BAT
      ├─git_STATUS+PULL+ADD+COMMIT+PUSH.BAT
      ├─MSYS_N02________________________________________________________________MSYS.txt
      ├─Project_MSYS.sublime-project
      ├─Project_MSYS.sublime-workspace
      ├─README.md
      ├─ZIP_7ZIP_________________________________________________________________ZZ.txt
      ├─ZIP_7ZIP_INCREMENTALE_{CRIPTA+DECODIFICA_INTERNO}_CASA_C.BAT
      ├─ZIP_7ZIP_INCREMENTALE_{CRIPTA+DECODIFICA_INTERNO}_SAN_GIORGIO_C.BAT
      ├─ZIP_7ZIP_INCREMENTALE_{CRIPTA+DECODIFICA_INTERNO}_UFFICIO_C.BAT
      ├─Zip_N70_CRIPTA.bat
      ├─Zip_N70_DECODIFICA.bat
      ├─Zip_N70_ElencoFile.txt
      ├─Zip_N70_ElencoFileESCLUSI.txt
      └─ZZ________________________________________________________________________ZZ.txt
   

# MSYS_OGGETTI _(ANALISI PRINCIPALE)


@ UTILIZZO GIORNALIER0
	dove il github:           git : icivixwalter/MSYS_OGGETTI
	COMANDO PLURIMO               : 
               git pull & git status & git add . & git commit -m "agg casa" & git push & git pull



# MODIFICHE
   2025_05_12_EFFICIENTATO_LA GESTIONE DEI FILE 
      nel progetto GEFILE_Frm01_GESTIONI ho efficientato la gestione dei file, sulla 
         PAG_01 = il button RECUPERA PATH  è attivo e salva nella casella di testo
                  Percorso_Txt

         PAG_02 =    IL COMANDO APRI TUTTI I FILE CON 3 PARAMETRI è stato
               sistemato ed ora è attivo come button e sia come doppio click sia 
               sul campo testo path che su file. 

   2025_07_25_MSYS_ATTIVA_GEST_OGGETTI = CREO FUNZIONE PER IMPORTAZIONE LINK TABELLE
      Note

            DATABASE DI PROVA
               il db di prova in cui testo la nuova procedur si trova in questa path:

                  START "apro DB TEST" c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\TMP_IMPORTA_MODULI\MSYS_MODELLO.mdb

                     @total@commader_(apro la path test con il totalcommander)
                  START "APRI PATH TEST" "C:\Program Files\totalcmd\TOTALCMD64.EXE" /L="c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\TMP_IMPORTA_MODULI\" /R="c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\TMP_IMPORTA_MODULI\"

            per la gestione del database MSYS_ATTIVA_GEST_OGGETTI.mdb e delle sue
            form e query occorre gestione il collegamento di diverse tabelle.
            Il modulo che gestisce il link si chiama UTILITA_MsysDF14Mdl05_IMPORTAZIONE_SEMPLICE_DI_OGGETTI.
            Per ora viene utilizzata la routine di attivazione delle importazione denominata
            InizializzaCollectionTABELLE con la quale vengono gestite le seguenti importazioni
            o link codificate con :
               questo codice----> MSYS_ATTIVA_GEST_OGGETTI_Link_Tables.

            I link di seguito indicato sono rappresentati in gruppi di definizioni e quindi
            possono riguardare piu tabelle e db es. gruppo df13:
                  tbl.Add "Nome", "Msys_DF13_}----------------------------------------------------@"
                  "Msys_DF13_}----------------------------------------------------@"

            per attivare la procedura di collegamento chiamare la sub:
                  ImportaTabelle()


         GRUPPO LINK EFFETTUATI:
            01)_DF01
            02)_DF02 *     = correggi la path  ... OK CORRETTA
            04)_DF04
            05)_DF11
            06)_DF12
            07)_DF13
            08)_DF14

               RISULTATI DEI COLLEGAMENTI

                  Tabella 1: MSys_DF01_}----------------------------------------------------@
                  Tabella 2: MSys_DF01_COMANDI
                  Tabella 3: MSys_DF02_}----------------------------------------------------@
                  Tabella 4: MSys_DF02_TIPO_OGGETTO
                  Tabella 5: MSys_DF05_}----------------------------------------------------@
                  Tabella 6: MSys_DF05_ATTRIBUTI_TABELLA
                  Tabella 7: MSys_DF11_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab
                  Tabella 8: MSys_DF11_}----------------------------------------------------@
                  Tabella 9: Msys_DF11_PROGETTI
                  Tabella 10: Msys_DF12_}----------------------------------------------------@
                  Tabella 11: Msys_DF12_GE_ObjProgetti
                  Tabella 12: Msys_DF13_}----------------------------------------------------@
                  Tabella 13: Msys_DF13_DLL_PROGETTO
                  Tabella 14: Msys_DF14_}----------------------------------------------------@
                  Tabella 15: Msys_DF14_DLL_LIBRERIE




OGGETTI CREATI MODIFICATI O MIGLIORATI
   DB MSys_MODULI.mdb
      2025.08.05
         START "APRI CARTELLA" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\"
         Ho creato un modulo che esporta tutti gli oggetti in un file.txt
         per ora ESPORTA TUTTE LE QUERY del db corrente.
         Il modulo si chiama :
            ESPORTA_Mdl00_}------------------------------------------------@
            ESPORTA_Mdl01_TUTTI_GLI_OGGETTI_IN_FILE_LOG


   DB MSys_MODULI.mdb
      codice ---> 2025.08.05_MSys_MODULI.mdb_(vedi REAME per le modifiche e miglioramenti nel modulo)

         START "APRI CARTELLA" "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_MODULI\MDB\"

         Sul progetto ho fatto le seguenti modifiche e migliramenti e creazioni di oggetti:
         Ho creato 4 macro per la cancellazione e la importazione degli oggetti moduli e tabelle:
            UTILITA_MsysDF14Mcr05_ATTIVA_CancellaModuli
            UTILITA_MsysDF14Mcr05_ATTIVA_CancellaTabelle
            UTILITA_MsysDF14Mcr05_ATTIVA_ImportaTabelle

         Ho migliorato il modulo di importazione semplice:
            UTILITA_MsysDF14Mdl05_}----------------------------------------@
            UTILITA_MsysDF14Mdl05_IMPORTAZIONE_SEMPLICE_DI_OGGETTI

   DB MSys_MODULI.mdb 06.08.2025
         CREATO MODULO: ESPORTA_Mdl01_TUTTI_GLI_OGGETTI_IN_FILE_HTML

            Sul db moduli ho fatto questi miglioramenti:
               c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TABELLE\MDB\MSys_TABELLE.mdb
               sul modulo:
                  ESPORTA_Mdl01_TUTTI_GLI_OGGETTI_IN_FILE_HTML, ho creato
               tutte le funzione per l'esportazione degli oggetti, TABELLE, QUERY, FORM, MODULI
               REPORT in fomato HTML VEDI CODICE ----> @esporta@query@HTML
               Con la macro SALVA_Mcr01_TUTTI_GLI_OGGETTI_IN_HTML, salvo tutti gli oggetti del db
               in OBJECT in formato html e quindi vale come salvataggi visibili e ricopiabili e
               quindi ricostruibili come codice. UNICA NOTA DA FINIRE sono i REPORT dove salva
               i dati in formato report  ma non la classe del report; quindi da rivedere.
               ATTENZIONE L'ESPORTAZIONE DELLA FORM MASTER deve essere duplicata, esportata e cancellata
               perche quella originale essendo aperto non puo essere esportata.
         MODIFICA MODULI 2025.08.25
            l'ulteriore modifica effettuata sul db moduli riguarda la FORM
               UTILITA_MsysDF14Frm05_ATTIVA_FUNZIONI
            la form vie utilizzata per la gestione MSYS_ATTIVA_GEST_OGGETTI.mdb è cioè per 
            il db principale; ho creato nella stessa almeno 3 tab con questo significato:
               TAB_PAG_01 = 01) CANCELLA ED IMPORTA OGGETTI UTILITA
                  serve per importare gli oggetti necessari per il funzionamento della
                  form master MSYS_ATTIVA_GEST_OGGETTI.mdb
                  con 4 button importa e cancella.

                  
                  



          TODO: rimande da controllare se questa esportazione funzona:

               1) COLLEGARE QUESTE TABELLE
                  MSys_TAB_}-----------------------------------------------------@
                  MSys_TABELLE
                  MSys_TABELLE_TMP
               2) IMPORTARE LE QUERY
                  MSys_TABELLE_Qry01_01_SELECT_TUTTE
                  MSys_TABELLE_Qry01_01_SELECT_TUTTE_TMP
               3) IMPORTARE LA FORM
                  MSys_TABELLE_Frm01_}-------------------------------------------@
                  MSys_TABELLE_Frm01_S01_TUTTE
                  MSys_TABELLE_Frm01_S01_TUTTE_TMP


