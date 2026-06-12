# PROGETTO TB99_LINK

## NOTE
con il file TB99_LINK_README.MD gestisco il tutti i collegamenti tra
i sub progetti di MSYS


## ATTIVITA
il sub progetto TB99_LINK acquisisce trami modulo i seguenti oggetti:



1) TABELLE LINK  = tutti i collegamenti alla tabelle dei vari progetti
2) ....TUTTI gli oggetti che apprtengono ai progetti delle tabelle  
       collegate... todo da fare??
         come importare solo i NOMI delle query, le Form , moduli dei progetti??
         Da salvar in una sotto tabella link?
         Msys_TB99_LINK_Object?

## ATTIVITA
LE TABELLE LINK DI TUTTI I SUB PROGETTI
Le tabelle link recuperate in tutti i mini progetti vengno salvate sulla tabella
master:  
   Msys_TB99_LINK
Alla tabella master si associa una DETAIL
   Msys_TB99_LINK_Object
In cui vengono salvati TUTTI GLI OGGETTI che appartengono alla singola tabella LINK
verificata.

Una TABELLA LINK verificata RAPPRESENTA COME INSIEME tutti gli oggetti che si trovano nel
database di origine della tabella link, es: Query, Report, Form ecc per cui i nomi degli 
oggetti apparteneti a tale gruppo vengono salvati nell TABELLA DETAIL con il codice di 
progetto di appartenenza.
SCOPO DELLA TABELLA DETAIL è quella di essere utilizzata per esportare ed importare gli oggetti
che appartengono all'insieme del PROGETTO agli altri progetti.

## DEFINIZIONE TIPO DI ESPORTAZIONI ED IMPORTAZIONI
ESPORTAZIONI
Ogni tabella collegata rappresenta 2 ATTIVITA: una attivita di Esportazione e sia una attivita di 
importazione. Per le esportazioni puoi : attivare una 
      - esportazione di DATI (sono i dati di backup salvati in un db esterno per ogni tabella collegata); 
      - esportazione di OGGETTI (query, form ecc. del mini progetto della tabella collegata);
IMPORTAZIONI
Per le importazioni si intende solo l'importazione DEI NOMI DEGLI OGGETTI 
e che vengono salvati nella tabella collegata denominata: Msys_TB99_LINK_Object
tramite il codice progetto; VENGONO SALVATI SOLO I NOMI DEGLI OGGETTI. In questo modo 
con  il collegamento delle due tabelle
   Msys_TB99_LINK
      |
      |key = COD_PROGETTO_s
      |
   Msys_TB99_LINK_Object
linkate in altri mini progetti puoi importare gli oggetti collegati a singole tabelle
verficate con precisione e velocita.


## COSTRUZIONE DELLA ESPORTAZIONE DATI
La esportazione dei dati della singola tabella collegata avviene mediante la query di 
estrazione delle tabelle collegata sia :
   VERIFICATE (collegamento attivo)
   SCELTE = true ossia segnalate per la esportazione dei dati.

   CREATO UNA ESPORTAZIONE 
      A) ESPORTAZIONE DATI DALLA TABELLA COLLEGATA AL PROGETTO DI BACKUP MSys_TB99_LINK_Backup.mdb


TODO: definire le modalita di esportazione ecc.... POSSIAMO DEFINIRE 
      IL TIPO DI ESPORTAZIONE:
      B) ESPORTARE GLI OGGETTI che appartengno alla singola tabella in altro db

      TIPO DI IMPORTAZIONE
      A) IMPORTARE GLI OGGETTI CHE APPARTENGONO ALLA SINGOLA TABELLA COLLEGATA
         al contenitore di oggetti : Msys_TB99_LINK_Object
         La tabella collegata appartiene ad un progetto, per cui ogni oggetto 
         importato in TB99_LINK_Obiect  appartiena al progetto.





## OBJECT TB99
tra gli oggetti ci sono 
   TABELLE
      Msys_TB99_LINK                   = Questa tabella ricostruisce per intero i

   QUERY_SELECT
      Msys_TB99_LINK_Qry99_01_SELECT_Archivio                        = estraggo archivio tb99
   QUERY INSERT
      Msys_TB99_LINK_Qry99_70_INSERT_ArchivioTabelleLINK             = inserisco da MsyObject tutte le tabelle link in tb99
