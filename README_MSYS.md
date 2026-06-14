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
            │__GE_MSys_DF01.mdb
                  │__MSys_DF01.mdb
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



   ALBERO CARTELLE E FILE  - @albero@progetto_(il Progetto master ed i sub progetti)_@albero@cartelle
         MSYS/
            │
            ├─GIT_BAT/
            ├─MDB/
            │  │
            │  │
            │  │----MSYS_ATTIVA_GEST_OGGETTI/             CODICE=            
            │  --------------------------------------------------------------------------------------------------------------------------
            │           PROJECT                          CODICE_PROJECT                    ATTIVITA
            │  --------------------------------------------------------------------------------------------------------------------------
            │            │___MSYS_BASE/                            │  CODICE =BASE    │    = @DB@BASE da utilzzare per tutti i mini progetti 
            │            │___MSys_DF01/                            │  CODICE =DF01    │    = DEFINIZIONE DEI COMANDI DEL MENU                         - @comandi@menu_(di tutto il @superprogetto) 
            │            │___MSys_DF02/                            │  CODICE =DF02    │    = DEFINIZIONE DEI TIPI DI OGGETTO                          - @TIPO@DI@OGGETTO_(vengono definiti i tipi di oggetto gesti, FORM, QUERY TABELLE ECC..) 
            │            │___MSys_DF05/                            │  CODICE =DF05    │    = DEFINIZIONE DELLE QUOTE DEGLI OGGETTI                    - @QUOTE@DI@PROGETTO_(con le quote per ogni subprogetto puoi definire i gruppi di oggetti utilizzati) TODO: fose inutile dupelicato controllare o sostituire con CODICE PROGETTO 
            │            │___MSys_DF10/                            │  CODICE =DF10    │    = DEFINIZIONE DEI DATABASE UTILIZZATI                      - @DATABASE@UTILIZZATI_(i vari database utilizzati per la gestione del progetto di importazone ed esportazione) todo: COME UTILIZZARLI? 
            │            │___MSys_DF11/                            │  CODICE =DF11    │    = DEFINIZIONE DEI CODICI DI PROGETTO                       - @CODICI@PROGETTO_(con la TABELLA @DF11 ad ogni sub progetto viene attribuito un @codice@progetto, anche al MASTER) 
            │            │___MSys_DF12/                            │  CODICE =DF12    │    = DEFINIZIONE DEGLI OGGETTI NECESSARI AL PROGETTO          - 8 oggetti necessari per oggi progetto @oggetti@object_(per ogni progetto occorrono 8 oggetti, QURY, FORM REPORT ECC.) TODO: come utilizzarli? 
            │            │___MSys_DF13/                            │  CODICE =DF13    │    = DEFINIZIONI DELL LIBRERIE DLL DEL PROGETTO               - @DLL_(per ogni progetto occorrono le librerie DLL da stabilire) 
            │            │___MSys_DF14/                            │  CODICE =DF14    │    = PATH DI SALVATAGGIO DELLE LIBRERIE ACCESS                - dove sono state salve @librerie_(dove si trovano o dove sono state salvate) 
            │            │___MSys_ELENCO/                          │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_ESPORTA/                         │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_FORMS/                           │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_FORMS_MASTER/                    │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_GE_FILE/                         │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_MACRO/                           │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSYS_MENU/                            │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_MODULI/                          │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSYS_OBJECT_DB_EST/                   │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSYS_OGGETTI_OLD/                     │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_PAGINE/                          │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_QUERY/                           │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_REPORT/                          │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_TABELLE/                         │  CODICE =123456  │    = da definire l'utilizzo??
            │            │___MSys_TB01_PROGETTI/                   │  CODICE=         │    = I CODICE DI PROGETTO E LA LORO DEFINIZIONE @CODICI@PROGETTI_(dove sono definiti i codice dei progetti la denominazione e cosa fanno)
            │            │___MSys_TB02_TIPO_OGGETTI/               │  CODICE=         │    = IL GRUPPO DI OGGETTI NECESSARI AL FUNZIONAMENTO DEI PROGETTI @gruppo@oggetti_(raggruppamento degli oggetti per ogni progetto)
            │            │___MSys_TB03_ELENCO_OGGETTI/             │  CODICE=         │    = ELENCO DEGLI OGGETTI ASSEGNATI ad ogni progetto @eleco@project_(tutti gli oggetti necessari)
            │            │___Msys_TB05_FORMS/                      │  CODICE=         │    = MSYS PER LA GESTIONE DI TUTTE LE FORMS *** IN RICOSTRUZIONE   
            │            │___Msys_TB40_PROG_MASTER/                │                  │    = LE FORM MASTER E LE SOTTOFORM MASTER  *** IN RIOSTRUZIONE
            │            │___MSys_TB41_VISUAL_PAG_MASTER/          │                  │    = LA VISUALIZZAZIONE DELLE PAGINE
            │            │___MSys_TUTTI_MODELLI/                   │                  │    = da studiare come ridefinire queti mini progetti
            │            │___TAB_CONTROL/                          │                  │    = da studiare come ridefinire queti mini progetti
            │            │___TAB_CONTROL_MASTER_(401_402_501_502)/ │                  │    = da studiare come ridefinire queti mini progetti
            │            │___TMP_IMPORTA_MODULI/                   │                  │    = da studiare come ridefinire queti mini progetti
            │            │___TMP_IMPORTA_MODULIOBJECT/             │                  │    = da studiare come ridefinire queti mini progetti
            │            │___TMP_OGGETTI_DA_CANCELLARE/            │                  │    = da studiare come ridefinire queti mini progetti
            │            │___TREVIEW/                              │                  │    = da studiare come ridefinire queti mini progetti
            │            │___UTIL/                                 │                  │    = da studiare come ridefinire queti mini progetti
            │
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
      

# HELP IN HTML
## COME COSTRUIRE UN HELP CON LA PAGINA HTML
codice -----> @attivo@pagina@htlm
occorre inserire un un campo l'evento KeyDown o @Tasto@giu e si attiva l'evento
per chiamare dall'esterno la pagina html


   '//==================================================================================================//
   '//             EVENTI CASELLA DI TESTO @MODELLO@HELP@HTML_(EVENTI CON KEY DOWN  )   *** INIZIO ***
   '//==================================================================================================//
   '//EVENTI: _
       '//01_?                             = ...........
       '//02_MODELLO_HELP_s_TXT_KeyDown     = TASTO GIU CON KEY DOWN
       '//'
       '//codice -----> @attivo@pagina@htlm_(chiamo un evento tasto giu ed attivo la @pagina@html)'
       '//              puoi attivarlo anche per click in tal caso devi disattivare gli if e Keycode'
   
   
   
   '//@HELP@HTML@MODELLO
   Private Sub MODELLO_HELP_s_TXT_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   
       
       ' Se l'utente preme F1 (codice 112)
       If KeyCode = vbKeyF1 Then  ' vbKeyF1 è una costante VBA = 112
           ' Percorso assoluto del file HTML
           Dim percorsoHelp As String
           Dim FileHelp As String
           Dim HelpHtml_s  As String
           '//il percorso help + il file help
           percorsoHelp = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_TB03_ELENCO_OGGETTI\HELP\FORM\PROGETTI_Msys_TB02Frm01_01_TIPO_OGGETTI\"
           FileHelp = "HELP_MODELLO_HELP_s_TXT.html"
           HelpHtml_s = percorsoHelp & FileHelp
           
           ' Apri il file HTML
           Application.FollowHyperlink HelpHtml_s
           
           ' Annulla l'azione predefinita di F1
           KeyCode = 0
       End If
   
   
   
   End Sub
   
   
   '//==================================================================================================//










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


