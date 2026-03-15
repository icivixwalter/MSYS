# TABELLA DI ESPORTAZIONI DEGLI OGGETTI IN UN DB ESTERNO

## DENOMINAZIONE
La tabella che definisce la path ed il progetto .md dove trasferire gli oggetti
si chiama :PROGETTI_MSYS_TB04_DB_ESTERNO



## CHIAVI
Le chiavi della tabella di esportazione degli oggetti in un db esterno è formata da
4 campi:
   - COD_PROGETTO_s                     = il codice progetto
   - TIPOGGETTO_s                       = Il tipo di oggetto da esportare
   - NOMEOGGETTO_s                      = Il nome dell'oggetto da esportare
   - DATABASE_DEST_s                    = Il database di destinazione che deve
                                          essere univoca per evitare di esportare piu volte
                                          l'oggetto sullo stesso database.
                                          Il database duplicato non è ammesso.


## CODICE_ESTRAZIONE_ACCESS
per estrare tutti gli oggetti che appartengono alla gestione della tabella 04 DB ESTERNO
i codice sono i seguenti:
   TB04_DB_ESTERNO                        = ESTRAI TUTTI GLI OGGETTI ACCESS
      TB04_DB_ESTERNO_Qry01_01            = ESTRAI TUTTI GLI OGGETTI QUERY che gestiscono la tabella TB04



### QUERY PER CODICE
      PROGETTI_Msys_TB04Qry01_}--------------------------------------@
   PROGETTI_Msys_TB04Qry01_01_Archivio                         = archivio db di destinazione
