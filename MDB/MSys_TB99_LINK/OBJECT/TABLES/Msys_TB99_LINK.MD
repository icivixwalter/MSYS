# TABELLA TB99 LINK

## NOTE
con il file Msys_TB99_LINK.MD rappresento la tabella tb99 che
ricostruisce uno CATALOGO DI SISTEMA uguale alla tabella di sistema
MSysObjects riportando tutti i campi della tabella di sistema.




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
            
            proprieta della tabella = CODICE;DESCRIZIONE;OK;Valido e raggiungibile;NOT_REGISTERED;Non censito in TB99;MISSING;Censito ma file/oggetto non trovato;INVALID_PATH;Path strutturalmente errato;NOT_CHECKED;Non ancora verificato

