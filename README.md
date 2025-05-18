# MSYS_OGGETTI



	git : icivixwalter/MSYS_OGGETTI

	git add .


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

               