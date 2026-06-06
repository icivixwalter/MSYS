# @PROGETTI@IN@LAVORAZIONE
   2025.10.12 = IL PROGETTO IN LAVORAZIONE è quello esporta ed importa oggetti, vedi codice
                @ESPORTA@IMPORTA@OGGETTI

# SCHEMA CARTELLE E ASPETTI GENERALI DI UN MULTIPROGETTO ACCESS
   Note
      schema delle cartelle e dei file sul percorso del progetto:
            c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\
      Un PROGETTO ACCESS è composto da varie parti divisi fisicamente in cartelle.
      La cartella principale e le cartelle secondarie che dividono il progetto in questo modo:
            PROGETTO MASTER
               Contiene il database MULTIPROGETTO denominati
                  MULTIPROGETTO_MASTER.MDB
            MINI PROGETTI
               GESTIONE_PROGETTO_MINI.mdb
               ProgettiDefinizioni.mdb
               ProgettiArchivio.mdb
      Esempio di questa divisione puo essere questa:
         MSYS_ATTIVA_GEST_OGGETTI.mdb
            |__GE_MSys_DF01.mdb
                  |__MSys_DF01.mdb
   Vantaggi
      con questa divisione si possono costruire i proggetti in modo separato e modulare
      e con prove di gestione sia nel mini progetto che nel progetto master, essendo
      facilmente manutenibili.
   Accortezze
      Nel mini progetto BASE.MDB viene costruito un GESTORE DI OGGETTI che ha come unica
      funzione quella di IMPORTARE GLI OGGETTI direttamente dai mini progetti escluso dal
      proprio progetto sorgente.
      Quindi si hanno 2 tabelle di gestione per questa attivita:
         1) TABELLA OGGETTI     = che contiene path+database e gli oggetti che devono essere importati
         2) TABELLA SORGENTI    = contiene i database archivio e database definizioni che non possono 
         3)                         attivare l'importazione degli oggetti per evitare le cancellazioni
                                    inavvertite.



   ALBERO CARTELLE E FILE
         MSYS/
            │
            ├─GIT_BAT/
            ├─MDB/
            |  |
            |  |
            |  |--------MSYS_ATTIVA_GEST_OGGETTI/        .... il progetto MASTER
            |                 |___MSYS_BASE/             .... tutti i mini progetti
            |                 |___MSys_DF01/
            |                 |___MSys_DF02/
            |                 |___MSys_DF05/
            |                 |___MSys_DF10/
            |                 |___MSys_DF11/
            |                 |___MSys_DF12/
            |                 |___MSys_DF13/
            |                 |___MSys_DF14/
            |                 |___MSys_ELENCO/
            |                 |___MSys_ESPORTA/
            |                 |___MSys_FORMS/
            |                 |___MSys_FORMS_MASTER/
            |                 |___MSys_GE_FILE/
            |                 |___MSys_MACRO/
            |                 |___MSYS_MENU/
            |                 |___MSys_MODULI/
            |                 |___MSYS_OBJECT_DB_EST/
            |                 |___MSYS_OGGETTI_OLD/
            |                 |___MSys_PAGINE/
            |                 |___MSys_QUERY/
            |                 |___MSys_REPORT/
            |                 |___MSys_TABELLE/
            |                 |___MSys_TB01_PROGETTI/
            |                 |___MSys_TB02_TIPO_OGGETTI/
            |                 |___MSys_TB03_ELENCO_OGGETTI/
            |                 |___Msys_TB05_FORMS/                         = MSYS PER LA GESTIONE DI TUTTE LE FORMS *** IN RICOSTRUZIONE   
            |                 |___Msys_TB40_PROG_MASTER/                   = LE FORM MASTER E LE SOTTOFORM MASTER  *** IN RIOSTRUZIONE
            |                 |___MSys_TB41_VISUAL_PAG_MASTER/             = LA VISUALIZZAZIONE DELLE PAGINE
            |                 |___MSys_TUTTI_MODELLI/
            |                 |___TAB_CONTROL/
            |                 |___TAB_CONTROL_MASTER_(401_402_501_502)/
            |                 |___TMP_IMPORTA_MODULI/
            |                 |___TMP_IMPORTA_MODULIOBJECT/
            |                 |___TMP_OGGETTI_DA_CANCELLARE/
            |                 |___TREVIEW/
            |                 |___UTIL/
            |
            ├─OBJECT/
            ├─XLS/
            ├─.gitignore
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

   ALBERO PROGETTO MSYS
      
# RICOSTRUIRE LA REPOSITORY LOCALE ALLINEATA ALLA REMOTA
   Note
      devi svuotare completamente la cartella MSYS anche con i dati della
      repository locale .git ed applicare i seguenti comandi in seguenza.
      Nota bene è valita solo per la repository REMOTA MSYS.git

   RICOSTRUIRE LA REPOSITORY LOCALE DI GIT 
         per la repository REMOTA:  https://github.com/icivixwalter/MSYS.git
         
         Tutti i passaggi per ricostruire la repository locale allineata con quella remota :
         
        01) git init → creata la repository locale vuota.
        02) git remote add origin ... → collegato il repository remoto.
        03) git remote -v → confermato che il remote è registrato correttamente.
        04) git fetch origin → scaricati tutti i branch e commit remoti senza modificare il tuo branch locale.
        05) git checkout -b calcolo origin/calcolo → creato il branch locale calcolo e impostato il tracking con il remoto.
        06) git pull → niente da aggiornare perché ora il tuo branch locale è già allineato con origin/calcolo.
      
      Ora la tua repository locale è perfettamente sincronizzata con il remoto e pronta per lavorare.
         
         TUTTI I COMANDI DA ESEGUIRE:
         
         
            cd c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS
            git init
            git remote add origin https://github.com/icivixwalter/MSYS.git
            git remote -v
            git fetch origin
            git checkout -b calcolo origin/calcolo
            git pull


# RICOSTRUIRE LA PATH MSYS CON IL FILE .BAT
   Note
      la path msys puo essere ricostruita con il comando dos di seguito indicato.
      Essa ricalca lo schema delle cartelle sopra indito:

   DOS

         @REM FILE ----> BAT_RICREA_PATH_MSYS.BAT
            @echo off
            setlocal enableextensions
            
            rem Percorso completo da ricostruire
            set "TARGET=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS"
            
            if exist "%TARGET%" (
                echo Esiste già: "%TARGET%"
            ) else (
                echo Creo: "%TARGET%"
                mkdir "%TARGET%"
                if errorlevel 1 (
                    echo ERRORE: impossibile creare "%TARGET%".
                    exit /b 1
                ) else (
                    echo Creato con successo.
                )
            )
            
            endlocal



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



## SHADOW SYSTEM CATALOG (CATALOGO SPECCHIO DEL SISTEMA)

CODICE: @SHADOW_SYSTEM_CATALOG
STRUTTURA
   
   MDB/
   │
   └── MSYS_ATTIVA_GEST_OGGETTI/            .... PROGETTO MASTER
       │
       ├── MSYS_BASE/
       ├── MSys_DF01/
       ├── MSys_DF02/
       ├── MSys_DF03/
       ├── MSys_DF05/
       ├── MSys_DF10/
       ├── MSys_DF11/
       ├── MSys_DF12/
       ├── MSys_DF13/
       ├── MSys_DF14/
       ├── MSys_ELENCO/
       ├── MSys_ESPORTA_IMPORTA/
       ├── MSys_FORMS/
       ├── MSys_FORMS_MASTER/
       ├── MSys_GE_FILE/
       ├── MSys_MACRO/
       ├── MSYS_MENU/
       ├── MSys_MODULI/
       ├── MSYS_OBJECT_DB_EST/
       ├── MSYS_OGGETTI_OLD/
       ├── MSys_PAGINE/
       ├── MSys_QUERY/
       ├── MSys_REPORT/
       ├── MSys_TABELLE/
       ├── MSys_TB01_PROGETTI/
       ├── MSys_TB02_TIPO_OGGETTI/
       ├── MSys_TB03_ELENCO_OGGETTI/
       ├── MSys_TB04_DB_ESTERNO/
       ├── Msys_TB05_FORMS/                 .... gestione FORMS
       ├── Msys_TB06_MACROS/
       ├── Msys_TB07_MENU/
       ├── Msys_TB08_MODULES/
       ├── Msys_TB09_QUERY/
       ├── Msys_TB10_REPORTS/
       ├── Msys_TB11_TABLES/
       ├── Msys_TB30_UTIL/
       ├── Msys_TB40_PROG_MASTER/           .... FORMS MASTER
       ├── MSys_TB41_VISUAL_PAG_MASTER/     .... visualizzazione pagine
       ├── MSys_TB99_LINK/
       ├── MSys_TUTTI_MODELLI/
       ├── MsysUTIL/
       ├── TAB_CONTROL/
       ├── TAB_CONTROL_MASTER_(401_402_501_502)/
       ├── TMP_ESPORTA_OGGETTI/
       ├── TMP_ESPORTA_OGGETTI_NUOVO_TEST/
       ├── TMP_IMPORTA_MODULI/
       ├── TMP_IMPORTA_MODULIOBJECT/
       ├── TMP_OGGETTI_DA_CANCELLARE/
       └── TREVIEW/
“SHADOW METADATA REPOSITORY”
cioè
SISTEMA REALE (MSysObjects)
        ↓ (replica controllata)
SISTEMA APPLICATIVO (MSys_TB99_LINK)

    MSysObjects (Access interno, non modificabile in sicurezza)
        ↓
MSys_TB99_LINK (mirror controllato e modificabile)
   Quindi hai costruito:
   un catalogo parallelo controllato dall’applicazione
   con un’aggiunta fondamentale:
      Verifica_b = stato di coerenza runtime dei link
STATUS DEL LINK DI SISTEMA
    il campo verifica deve essere trasformato in testo con queste opzioni:
    Verifica_t = nuovo campo verifica

    STATUS:

      - OK              → valido e raggiungibile
      - NOT_REGISTERED  → non censito in TB99
      - MISSING         → censito ma file/oggetto non trovato
      - INVALID_PATH    → path strutturalmente errato
      - NOT_CHECKED     → non ancora verificato
            
            proprieta della tabella = OK;Valido e raggiungibile;NOT_REGISTERED;Non censito in TB99;MISSING;Censito ma file/oggetto non trovato;INVALID_PATH;Path strutturalmente errato;NOT_CHECKED;Non ancora verificato;
