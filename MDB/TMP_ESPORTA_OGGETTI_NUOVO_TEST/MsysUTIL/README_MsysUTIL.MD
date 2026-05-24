# ANALISI ESPORTAZIONE OGGETTI UTILITA

## NOTE
   Con il file 
   README_TB05_ANALISI_ESPORTAZIONE
   gestisco la esportazione degli oggetti del progetto istruttore.
   
   Parto dalla tabella:
   
   PROGETTI_Msys_TB01_PROJECT                            = contiene IL CODICE PROGETTO la denominazione e la quota
      |
      |---key: COD_PROGETTO_s
      |
   PROGETTI_Msys_TB02_TIPO_OGGETTI                        = tabella collegata che contiene i tipi di oggetto da gestire, query, form report ecc..
      |
      |---key: PROGETTI_Msys_TB03_ELENCO_ObjProgetti
      |
   PROGETTI_Msys_TB03_ELENCO_ObjProgetti                  = tabella in cui defisicini e scegli gli oggetti necessari per ogni progetto
      |
      |---key: PROGETTI_Msys_TB03_ELENCO_ObjProgetti
      |
   PROGETTI_MSYS_TB04_DB_ESTERNO                           = tabella in cui specichi il progetto destinazione da 1 a N per OGNI OGGETTO



## OGGETTI PER L'ESPORTAZIONE
   Creo le query per esportare gli oggetti dalla path di partenza alla destinazione.
   Esempio se 


## TABELLA DI SISTEMA ACCESS MSYOBJECTS
### Tabella di sistema MSysObjects (Microsoft Access)

La tabella di sistema `MSysObjects` contiene il catalogo interno degli oggetti del database Access:

- Tabelle
- Query
- Maschere
- Report
- Moduli
- Macro
- Relazioni
- Oggetti temporanei
- Oggetti di sistema

---

### Struttura campi
|----------------|--------------------------|--------------------------------------|
| Campo          | Significato              | Note                                 |
|----------------|--------------------------|--------------------------------------|
| `Name`         | Nome dell'oggetto        | Nome tabella, query, form, report    |
| `Type`         | Tipo oggetto             | Identifica il tipo oggetto Access    |
| `Id`           | ID interno univoco       | Identificatore interno Access        |
| `ParentId`     | ID oggetto padre         | Relazione con oggetti contenitori    |
| `Flags`        | Proprietà interne        | Bitmask oggetti nascosti/sistema     |
| `DateCreate`   | Data creazione           | Timestamp creazione oggetto          |
| `DateUpdate`   | Ultima modifica          | Timestamp ultima modifica            |
| `Owner`        | Proprietario oggetto     | Vecchia sicurezza Jet                |
| `Database`     | Database origine         | Collegamenti / repliche              |
| `Connect`      | Stringa connessione      | Tabelle collegate                    |
| `ForeignName`  | Nome remoto oggetto      | Nome tabella/query remota            |
| `Lv`           | Campo interno Access     | Uso interno non documentato          |
| `LvExtra`      | Campo interno Access     | Uso interno non documentato          |
| `LvModule`     | Campo interno Access     | Uso interno non documentato          |
| `LvProp`       | Campo interno Access     | Uso interno non documentato          |
| `RmtInfoLong`  | Informazioni replica     | Replica / remoto                     |
| `RmtInfoShort` | Informazioni replica     | Replica / remoto                     |
---

### Valori del campo Type

|-----------|------------------------------|
| Type      | Tipo oggetto                 |
|-----------|------------------------------|
| `1`       | Tabella locale               |
| `4`       | Tabella collegata ODBC       |
| `5`       | Query                        |
| `6`       | Tabella collegata non ODBC   |
| `-32768`  | Maschera (Form)              |
| `-32764`  | Report                       |
| `-32761`  | Modulo VBA                   |
| `-32756`  | Macro                        |
| `-32766`  | Query Pass-Through           |
| `-32758`  | Relazione                    |
| `-32757`  | Oggetto nascosto / sistema   |

---

### Significato del campo Flags

Il campo `Flags` è una bitmask.

### Significato del campo Flags

|--------|---------------------------------|
| Valore | Significato                     |
|--------|---------------------------------|
| `1`    | Oggetto di sistema              |
| `2`    | Oggetto nascosto                |
| `3`    | Oggetto di sistema + nascosto   |

Esempio:

