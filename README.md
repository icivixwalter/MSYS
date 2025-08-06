# MSYS_OGGETTI


@ UTILIZZO GIORNALIER0
	dove il github:           git : icivixwalter/MSYS_OGGETTI
	COMANDO PLURIMO               : 
               git pull & git status & git add . & git commit -m "agg casa" & git push & git pull


# GIT
      
      PROFILO GIT SE DA ERRORE
         
            P
               @PROFILO GIT
                  per i profilo git questa è l'utenza 
                  email: icivixwalter@gmail.com
                  passw: walter_01_
      
                  per configurare il profilo per tutto utilizzo globale con questi comandi
                     git config --global user.email "you@example.com"
                     git config --global user.name "Your Name"
      
                  se vuoi impostar il profilo solo per la  repository
                  corrente utilizza:
                  git config user.email "you@example.com"
                  git config user.name "Your Name"
      
                     esempio
                     git config user.email "icivixwalter@gmail.com"
                     git config user.name "walter"
      
            W
               warning LF
                  warning: in the working copy of 'Project_MSYS.sublime-workspace', LF will be replaced by CRLF the next time Git touches it
      
               Cosa significa il messaggio:
                  LF (Line Feed): È il carattere di fine riga standard su sistemi Unix-like (Linux, macOS).
                  CRLF (Carriage Return + Line Feed): È il carattere di fine riga standard su Windows.
                  Git cerca di gestire queste differenze automaticamente per evitare problemi tra i sistemi operativi.
      
               1 Come eliminare il warning:
                  Configurare Git per gestire i fine riga automaticamente: Puoi impostare una politica globale in Git per gestire i line endings e evitare che venga visualizzato il warning.
      
                  Esegui il comando:
                     git config --global core.autocrlf true
      
                  Questo farà in modo che Git convertirà automaticamente i fine riga CRLF in LF quando effettui il commit, e li convertirà di nuovo in CRLF quando fai il checkout su Windows. Questa configurazione è molto utile se lavori su un sistema Windows, ma collabora con altri che usano Linux/macOS.
      
      
               2 Alternative:
      
                     Se non vuoi che Git faccia nessuna conversione automatica dei fine riga e preferisci mantenere i line endings così come sono, puoi usare:
                        git config --global core.autocrlf false
      
      
                     Tuttavia, questa configurazione può portare a conflitti tra LF e CRLF se più sviluppatori usano diversi sistemi operativi.
      
      
                  Se vuoi forzare Git a usare solo LF (sempre uguali per tutti), puoi configurarlo con:
      
                        git config --global core.eol lf
      
      
                  Questo dirà a Git di usare sempre LF, ignorando le differenze tra CRLF e LF sui sistemi operativi.
      
               3 Aggiornare i file esistenti: 
                  Se il warning si riferisce a file già esistenti, come nel caso di Project_MSYS.sublime-workspace, potresti voler normalizzare i line endings per questi file esistenti. Puoi fare così:
      
                  a. Rimuovere il file dal controllo versione (senza eliminarlo dal disco):
      
                     git rm --cached Project_MSYS.sublime-workspace
      
                  b. Aggiungere di nuovo il file con il line ending corretto:
                     git add Project_MSYS.sublime-workspace
      
                  c. Eseguire il commit:
      
                     git commit -m "Normalizza i line endings del file Project_MSYS.sublime-workspace"
      
      
               In breve:
                  Se vuoi evitare il messaggio di warning e lasciare che Git gestisca automaticamente i fine riga, esegui:
      
                     git config --global core.autocrlf true
      
      
      
      CREARE I RAMI
         Gestione delle filiali
            Un singolo repository Git può mantenere più rami di sviluppo. Per creare 
            un nuovo ramo denominato , utilizzareexperimental
               git branch experimental
               git branch ramo2			= 2 esempio
      
            per controllare quanti rami esistono eseguire
               git branch
            Verrà visualizzato un elenco di tutti i rami esistenti:
      
            Io ho il main ma posso creare un ramo secondario di appoggio
               git branch appoggio
      
            ho creato il brach e mi trovo ancora su main con l'asterisco
            per passare al nuovo branch digitare
               git switch appoggio
      
      
            Ora modifica un file, esegui il commit del file modificare e tornare al ramo
               git commit -a
               git branch main
      
            A questo punto i due rami si sono divisi, con cambiamenti diversi fatto in ciascuno. Per unire le modifiche apportate in ,  
               git merge appoggio
      
            Se le modifiche non sono in conflitto, il gioco è fatto. In caso di conflitti, I marcatori verranno lasciati nei file problematici che mostrano il conflitto;
      
               git diff
      
      
            per veder tutte le differenze tra i git digitare
               gitk
      
      
      GIT_REMOTE
         per attivare il salvataggio all'inizio eseguire i due comandi:
             git remote add <name> <url> = 	git remote add "walter" "https://github.com/icivixwalter/MSYS.git"
                                                         
         and then push using the remote name
      
          git push <name> = git push --set-upstream origin calcolo
      
      
          se da questo errore:
          fatal: The current branch calcolo has no upstream branch.
         To push the current branch and set the remote as upstream, use
      
                eseguire il comando:T&
                   git push --set-upstream origin calcolo
      
      
      GIT_ADD
         per aggiornare la REPOSITORY occorre eseguire i seguenti comandi concatenati:
            
            @comando@unico@git
      
            git pull & git status & git add . & git commit -m "agg ufficio" & git push
      
      
      
         per escludere file e cartelle utilizzare il file:
      
         .gitignore
      
         senza estensione per le cartelle al suo interno 
         utilizzare il nome della cartella:
            AA_SALVATAGGI/
      
            per il file in sotto cartelle il percorso completo
            esempio: 
               c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\IMPEGNI\MDB\X_IMPEGNI_RICREA_OGGETTI.mdb
      
      
      GIT_ERRORI
         ERRORE DI fast-forward  (AVANTI VELOCE, INOLTRARE VELEOCE)
            :: Risolve il problema di git push con errori di fast-forward
            :: --rebase assicura che le modifiche locali vengano applicate sopra le modifiche remote
            :: senza creare un merge non necessario
            @echo Correggere errore di git push...
            git pull --rebase
            git push
      
         ERRORE_PULL_REBASE
      
               git pull --rebase
               error: cannot pull with rebase: You have unstaged changes.
               error: Please commit or stash them.
      
               L'errore indica che ci sono modifiche non salvate nel repository locale. Per risolverlo, hai alcune opzioni:
      
                  Opzione 1: Committare le modifiche LOCALI e poi fare il pull
                     Se vuoi mantenere le modifiche:
      
                     git add .
                     git commit -m "Salvataggio delle modifiche prima del pull"
                     git pull --rebase
      
      
                  Opzione 2: Salvare temporaneamente le modifiche con stash
      
                        Se vuoi applicare il pull senza committare subito le modifiche:
      
                        git stash
                        git pull --rebase
                        git stash pop  # Ripristina le modifiche salvate
      
      
                  Opzione 3: Scartare le modifiche locali (Attenzione: perderai i cambiamenti non salvati!)
      
                     Se non hai bisogno delle modifiche locali e vuoi forzare l'allineamento con il repository remoto:
      
      
                        git reset --hard
                        git pull --rebase
      


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


