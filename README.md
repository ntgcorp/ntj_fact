# ntj_fact — user manual / manuale d'uso

[English](#english) | [Italiano](#italiano)

<a id="english"></a>
## English

### Purpose
`ntj_fact.py` is a **ntJobsApp** that manipulates known data files in batch mode,
plus a **database command group** (`ntj_fact_db.py`, loaded only if used).
For a full explanation of what a ntJobsApp is (INI in, `.end` INI out, exit codes,
`acJobsApp` orchestration), see **https://github.com/ntgcorp/ntJobsApp**.

Scope: manipulate data automatically — **not** import/export interactively.
Data files: XLS/XLSX (Excel), CSV, JSON, XML, TXT, SQL.
Databases: `SQLITE`, `ACCESS`, `ORACLE`, `POSTGRES`, `MARIADB` (drivers:
stdlib `sqlite3`, `pyodbc`, `oracledb`, `psycopg`, `pymysql`).

### INI structure
```ini
[CONFIG]
TYPE=NTJOBS.APP.1
NAME=NTJ_FACT_MACH0_TEST
LOG=mach0_utenti_test
EXIT=TRUE

[JOB_01]
COMMAND=XLS.TAB.MERGE
...

[DB_ANAG]              ; only if DB commands are used: one section per DB, ID = ANAG
TYPE=SQLITE
PATH=K:\dati\anag.db
```
- Keys are uppercased on read; `$VARIABLES` from `[CONFIG]` are expanded
  in job values (`FILE.IN.SHEET=$YA54M0`, also `$ENV.NAME`).
- Reserved keys (`TS.START/TS.END/RETURN.TYPE/RETURN.VALUE/RETURN.FILE.*`)
  must never appear in the input INI; they are written to the `.end` file.

### File extensions used

| Ext. | Role | Commands |
|---|---|---|
| `.ini` | input job file | all (launch) |
| `.end` | result INI (same basename as input) | all (output) |
| `.txt` | list of input files, one path per line | `XLS.TAB.MERGE` (`FILE.IN.SHEETS.TXT`) |
| `.xls` / `.xlsx` (any case) | input and output workbooks | `XLS.TAB.MERGE`, `XLS.SHEET.APPEND`, `SQL.EXEC` (out), `IMPORT` (in) |
| `.csv` | `;`-delimited data, header row required, utf-8(-sig) | `IMPORT` (`FILE.IN.CSV`) |
| `.json` | table structure (`FILE.IN.STRUCT`) or array-of-objects data | `TAB.CREATE.JSON`, `DB.IMPORT` (`FILE.IN.JSON`) |
| `.sql` | single statement text | `SQL.EXEC` (`FILE.IN.SQL`) |
| `.db` / `.sqlite` | SQLite database file | `DB.*` (`PATH`) |
| `.mdb` / `.accdb` | Access database file | `DB.*` (`PATH`) |

### Commands — files

**`XLS.TAB.MERGE`** — merge several xls/xlsx files sharing the same header
into one result file with a single leading header.
- `FILE.IN.SHEETS.TXT`: txt file (must exist) with one input file path per line.
- `SHEET.OUT`: sheet name to read from inputs and write to output
  (matched case-insensitively, e.g. `UTENTI` finds `utenti`; output uses the requested name).
- `FILE.OUT.XLS`: output file, extension `.xls` or `.xlsx` (any case).
- Returns row count (`RETURN.VALUE`) + output file (`RETURN.FILE.XLS`).

**`XLS.SHEET.APPEND`** — copy sheet `SHEET.IN` from `FILE.IN.SHEET` into
`FILE.OUT.XLS`, which is first created as a fresh copy of `FILE.IN.XLS`
(overwrites any previous output; other sheets of the base are kept).
If the destination sheet already exists it is emptied first, then the
source content is pasted **values-only** (never formulas), plus merged cells
and row/column sizes. Destination sheet QueryTables are normalized to plain tables.
- `FILE.IN.XLS`: base file (copied to output).
- `FILE.IN.SHEET`: file the source sheet is read from (supports `$VARIABLES` from `[CONFIG]`).
- `SHEET.IN`: source sheet name (case-insensitive match).
- `SHEET.OUT`: destination sheet name, optional (default `SHEET.IN`); if it
  already exists it is emptied and refilled, never appended to.
- `FILE.OUT.XLS`: output file (`.xls` or `.xlsx`, any case).
- Returns copied-row count + output file.

### Commands — databases (`ntj_fact_db.py`, spec: `prompt_ntj_fact_db.md`)

Each DB is declared once as `[DB_<ID>]` (`TYPE`, `PATH` or `HOST`/`PORT`/…,
optional `TIMEOUT`/`READONLY`, `$ENV.` recommended for `PASSWORD`).
Open DBs stay open across jobs (multiple DBs can be open); any still open
at end of run are closed automatically. A command on a non-open DB fails
with `DB <id> non aperto (manca DB.OPEN)`.

**`DB.OPEN`** — `DB.ID=<id>`: validate fields for `TYPE`, check driver,
open and store `HANDLE`. Double OPEN is an error.

**`DB.CREATE`** — `DB.ID=<id>`: create the container if missing (SQLite file,
Access via ADOX/COM, `CREATE DATABASE` with admin rights for server DBs),
then stays closed. Already exists = INFO, not an error.

**`DB.CLOSE`** — `DB.ID=<id>`: close `HANDLE` (no implicit commit). Closing
a non-open DB is an error.

**`SQL.EXEC`** — run ONE statement: `DB.ID=<id>`, `SQL=<statement>` xor
`SQL.FILE=<file.sql>` (exactly one). DDL/DML/DCL/SELECT allowed; no
multi-statement. `SELECT` without `FILE.OUT.XLS` returns only the rowcount;
with `FILE.OUT.XLS` (`.xls`/`.xlsx`) writes the rows + `RETURN.FILE.XLS`.
Returns rows (`RETURN.VALUE`).

**`TAB.ZAP`** — empty a table: `DB.ID=<id>`, `TABLE=<tab>`. Always
`DELETE FROM`, never `TRUNCATE`. Missing table is an error.
Returns deleted rows.

**`TAB.CREATE.JSON`** — create a table from a JSON structure file:
`DB.ID=<id>`, `FILE.IN.STRUCT=<file.json>`, `TABLE` optional (default from JSON,
must match if given). Neutral types (`INTEGER/TEXT/DECIMAL/DATE/
DATETIME/BOOLEAN`) translated to DDL per `TYPE`. Existing table is an error.

**`DB.IMPORT`** — load CSV or JSON data with optional JSON structure file:
`DB.ID=<id>`, `FILE.IN.STRUCT=<file.json>` (optional: if given, the file must
comply with it), `TABLE` optional (default from JSON; required without STRUCT),
`FILE.IN.CSV=<f>` xor `FILE.IN.JSON=<f>`, `MODE=APPEND|REPLACE`
(default `APPEND`; `REPLACE` = zap first). CSV: `;` delimiter, header
required; columns by name (case-insensitive), extras ignored,
missing-nullable excluded, missing-not-null → error with row number.
Without STRUCT, columns are read from the existing table (all nullable).
Returns loaded rows.

**`DB.EXPORT`** — export a query to file: `DB.ID=<id>`, `OUT.QUERY=<select>`
(SELECT only), `FILE.OUT.QUERY=<f>`, `OUT.TYPE=CSV|JSON` (required; extension
must match). CSV: `;` delimiter, header; JSON: array of objects.
Returns rows + `RETURN.FILE.QUERY`.

### Usage
```cmd
REM run with CWD = folder of the INI file, e.g. mach0/
python ntj_fact.py <job.ini>
```

Example (`mach0/mach0_utenti_test.ini`):
```ini
[CONFIG]
TYPE=NTJOBS.APP.1
NAME=NTJ_FACT_MACH0_TEST
LOG=mach0_utenti_test
EXIT=TRUE

[JOB_01]
COMMAND=XLS.TAB.MERGE
FILE.IN.SHEETS.TXT=mach0_xls_utenti_test.txt
SHEET.OUT=UTENTI
FILE.OUT.XLS=MACH0_UTENTI.XLSX
```

Example append job:
```ini
[JOB_CMLTV]
COMMAND=XLS.SHEET.APPEND
FILE.IN.XLS=UTENTI_MACH0_CMLTV_TEMP.XLSX
FILE.IN.SHEET=$YA54M0
SHEET.IN=UTENTI
SHEET.OUT=YA54M0
FILE.OUT.XLS=UTENTI_MACH0_CMLTV.XLSX
```

Example DB jobs:
```ini
[DB_ANAG]
TYPE=SQLITE
PATH=K:\dati\anag.db

[JOB_OPEN]
COMMAND=DB.OPEN
DB.ID=ANAG

[JOB_LOAD]
COMMAND=DB.IMPORT
DB.ID=ANAG
FILE.IN.STRUCT=anag_utenti.json
FILE.IN.CSV=utenti.csv
MODE=REPLACE

[JOB_CLOSE]
COMMAND=DB.CLOSE
DB.ID=ANAG
```

- Result: `<job>.end` INI file next to the input INI; exit code `0` ok,
  `1` startup/INI error, `2` one or more jobs failed.
- Run with **CWD = folder of the INI** (`FILE.*` inputs are checked by basename in CWD;
  `FILE.OUT.*` outputs are skipped by the check).
- Dependencies: `pandas` + `openpyxl` (xlsx) + `xlwt` (xls writing) + `xlrd` (xls re-read).
  `.xls` output is written directly with `xlwt` (pandas 3.x no longer writes BIFF);
  uppercase extensions (`.XLS`, `.XLSX`) are accepted.
  DB drivers per `TYPE` (checked up front with explicit install message):
  `sqlite3` stdlib, `pyodbc`, `oracledb`, `psycopg`, `pymysql`.
- Specs: `Prompt_ntj_fact.md` (files, binding), `prompt_ntj_fact_db.md` (databases).
  Known issues/limits: `ntj_fact_problems.md`.

---

<a id="italiano"></a>
## Italiano

### Scopo
`ntj_fact.py` è una **ntJobsApp** che manipola in modalità batch file di dati "conosciuti",
più un **gruppo comandi database** (`ntj_fact_db.py`, caricato solo se usato).
Per la spiegazione di cosa fa una ntJobsApp (INI in ingresso, INI `.end` in uscita,
exit code, orchestrazione `acJobsApp`), vedi **https://github.com/ntgcorp/ntJobsApp**.

Scopo: manipolare automaticamente i dati — **non** importarli/esportarli a mano.
File di dati: XLS/XLSX (Excel), CSV, JSON, XML, TXT, SQL.
Database: `SQLITE`, `ACCESS`, `ORACLE`, `POSTGRES`, `MARIADB` (driver:
`sqlite3` stdlib, `pyodbc`, `oracledb`, `psycopg`, `pymysql`).

### Struttura INI
```ini
[CONFIG]
TYPE=NTJOBS.APP.1
NAME=NTJ_FACT_MACH0_TEST
LOG=mach0_utenti_test
EXIT=TRUE

[JOB_01]
COMMAND=XLS.TAB.MERGE
...

[DB_ANAG]              ; solo se si usano comandi DB: una sezione per DB, ID = ANAG
TYPE=SQLITE
PATH=K:\dati\anag.db
```
- Le chiavi sono convertite in maiuscolo in lettura; le `$VARIABILI` di `[CONFIG]`
  sono espanse nei valori dei job (`FILE.IN.SHEET=$YA54M0`, anche `$ENV.NOME`).
- Le chiavi riservate (`TS.START/TS.END/RETURN.TYPE/RETURN.VALUE/RETURN.FILE.*`)
  non devono mai comparire nell'INI di ingresso; sono scritte nel file `.end`.

### Estensioni file usate

| Est. | Ruolo | Comandi |
|---|---|---|
| `.ini` | file job di ingresso | tutti (lancio) |
| `.end` | INI di risultato (stesso basename dell'ingresso) | tutti (output) |
| `.txt` | lista file di ingresso, un percorso per riga | `XLS.TAB.MERGE` (`FILE.IN.SHEETS.TXT`) |
| `.xls` / `.xlsx` (qualsiasi case) | workbook di ingresso e uscita | `XLS.TAB.MERGE`, `XLS.SHEET.APPEND`, `SQL.EXEC` (out), `IMPORT` (in) |
| `.csv` | dati delimitati `;`, header obbligatorio, utf-8(-sig) | `IMPORT` (`FILE.IN.CSV`) |
| `.json` | struttura tabella (`FILE.IN.STRUCT`) o dati array-di-oggetti | `TAB.CREATE.JSON`, `DB.IMPORT` (`FILE.IN.JSON`) |
| `.sql` | testo singolo statement | `SQL.EXEC` (`FILE.IN.SQL`) |
| `.db` / `.sqlite` | file database SQLite | `DB.*` (`PATH`) |
| `.mdb` / `.accdb` | file database Access | `DB.*` (`PATH`) |

### Comandi — file

**`XLS.TAB.MERGE`** — unisce più file xls/xlsx con identico header in un unico
file di risultato con un solo header iniziale.
- `FILE.IN.SHEETS.TXT`: file txt (deve esistere) con un percorso di file per riga.
- `SHEET.OUT`: foglio da leggere negli input e scrivere in output
  (confronto case-insensitive, es. `UTENTI` trova `utenti`; in scrittura si usa il nome richiesto).
- `FILE.OUT.XLS`: file di output, estensione `.xls` o `.xlsx` (maiuscola o minuscola).
- Ritorna il numero di righe scritte (`RETURN.VALUE`) + file di output (`RETURN.FILE.XLS`).

**`XLS.SHEET.APPEND`** — copia il foglio `SHEET.IN` dal file `FILE.IN.SHEET`
nel file `FILE.OUT.XLS`, creato prima come copia fresca di `FILE.IN.XLS`
(sovrascrive l'output precedente; gli altri fogli della base restano invariati).
Se il foglio di destinazione esiste già viene prima svuotato, poi il contenuto
sorgente è incollato **solo valori** (mai formule), più celle unite
e dimensioni righe/colonne. Le QueryTable di destinazione sono normalizzate
a tabelle semplici.
- `FILE.IN.XLS`: file base (copiato in output).
- `FILE.IN.SHEET`: file da cui si legge il foglio sorgente (supporta `$VARIABILI` di `[CONFIG]`).
- `SHEET.IN`: nome del foglio sorgente (confronto case-insensitive).
- `SHEET.OUT`: nome del foglio destinazione, opzionale (default `SHEET.IN`); se
  esiste già viene svuotato e riempito, mai accodato.
- `FILE.OUT.XLS`: file di output (`.xls` o `.xlsx`, maiuscolo o minuscolo).
- Ritorna il numero di righe copiate + file di output.

### Comandi — database (`ntj_fact_db.py`, specifica: `prompt_ntj_fact_db.md`)

Ogni DB è dichiarato una volta come `[DB_<ID>]` (`TYPE`, `PATH` oppure
`HOST`/`PORT`/…, opzionali `TIMEOUT`/`READONLY`, `$ENV.` raccomandato per
`PASSWORD`). I DB aperti restano aperti tra i job (più DB aperti insieme);
quelli ancora aperti a fine run sono chiusi in automatico. Un comando su
DB non aperto fallisce con `DB <id> non aperto (manca DB.OPEN)`.

**`DB.OPEN`** — `DB.ID=<id>`: valida i campi per `TYPE`, verifica il driver,
apre e salva `HANDLE`. Doppio OPEN è errore.

**`DB.CREATE`** — `DB.ID=<id>`: crea il contenitore se manca (file SQLite,
Access via ADOX/COM, `CREATE DATABASE` con diritti admin per i DB server),
poi resta chiuso. Esiste già = INFO, non errore.

**`DB.CLOSE`** — `DB.ID=<id>`: chiude `HANDLE` (senza commit implicito).
Chiudere un DB non aperto è errore.

**`SQL.EXEC`** — esegue UN solo statement: `DB.ID=<id>`, `SQL=<statement>` xor
`SQL.FILE=<file.sql>` (esattamente uno). Ammessi DDL/DML/DCL/SELECT; niente
multi-statement. `SELECT` senza `FILE.OUT.XLS` ritorna solo il conteggio;
con `FILE.OUT.XLS` (`.xls`/`.xlsx`) scrive le righe + `RETURN.FILE.XLS`.
Ritorna le righe (`RETURN.VALUE`).

**`TAB.ZAP`** — svuota una tabella: `DB.ID=<id>`, `TABLE=<tab>`. Sempre
`DELETE FROM`, mai `TRUNCATE`. Tabella mancante è errore.
Ritorna le righe cancellate.

**`TAB.CREATE.JSON`** — crea tabella da file JSON di struttura:
`DB.ID=<id>`, `FILE.IN.STRUCT=<file.json>`, `TABLE` opzionale (default da JSON,
se dato deve coincidere). Tipi neutrali (`INTEGER/TEXT/DECIMAL/DATE/
DATETIME/BOOLEAN`) tradotti in DDL per `TYPE`. Tabella esistente è errore.

**`DB.IMPORT`** — carica dati da CSV o JSON con JSON di struttura facoltativo:
`DB.ID=<id>`, `FILE.IN.STRUCT=<file.json>` (opzionale: se dato, il file deve
rispettarlo), `TABLE` opzionale (default da JSON; obbligatorio senza STRUCT),
`FILE.IN.CSV=<f>` xor `FILE.IN.JSON=<f>`, `MODE=APPEND|REPLACE`
(default `APPEND`; `REPLACE` = zap prima). CSV: delimitatore `;`, header
obbligatorio; colonne per nome (case-insensitive), extra ignorate,
mancanti-nullable escluse, mancanti-notnull → errore con riga. Senza STRUCT,
colonne lette dalla tabella esistente (tutte nullable). Ritorna le righe caricate.

**`DB.EXPORT`** — esporta una query in file: `DB.ID=<id>`, `OUT.QUERY=<select>`
(solo `SELECT`), `FILE.OUT.QUERY=<f>`, `OUT.TYPE=CSV|JSON` (obbligatorio;
estensione deve coincidere). CSV: delimitatore `;`, header; JSON: array di
oggetti. Ritorna righe + `RETURN.FILE.QUERY`.

### Uso
```cmd
REM lanciare con CWD = cartella del file INI, es. mach0/
python ntj_fact.py <job.ini>
```

Esempio (`mach0/mach0_utenti_test.ini`):
```ini
[CONFIG]
TYPE=NTJOBS.APP.1
NAME=NTJ_FACT_MACH0_TEST
LOG=mach0_utenti_test
EXIT=TRUE

[JOB_01]
COMMAND=XLS.TAB.MERGE
FILE.IN.SHEETS.TXT=mach0_xls_utenti_test.txt
SHEET.OUT=UTENTI
FILE.OUT.XLS=MACH0_UTENTI.XLSX
```

Esempio job di append:
```ini
[JOB_CMLTV]
COMMAND=XLS.SHEET.APPEND
FILE.IN.XLS=UTENTI_MACH0_CMLTV_TEMP.XLSX
FILE.IN.SHEET=$YA54M0
SHEET.IN=UTENTI
SHEET.OUT=YA54M0
FILE.OUT.XLS=UTENTI_MACH0_CMLTV.XLSX
```

Esempio job DB:
```ini
[DB_ANAG]
TYPE=SQLITE
PATH=K:\dati\anag.db

[JOB_OPEN]
COMMAND=DB.OPEN
DB.ID=ANAG

[JOB_LOAD]
COMMAND=DB.IMPORT
DB.ID=ANAG
FILE.IN.STRUCT=anag_utenti.json
FILE.IN.CSV=utenti.csv
MODE=REPLACE

[JOB_CLOSE]
COMMAND=DB.CLOSE
DB.ID=ANAG
```

- Risultato: file INI `<job>.end` accanto all'INI di ingresso; exit code `0` tutto ok,
  `1` errore di avvio/INI, `2` uno o più job in errore.
- Lanciare con **CWD = cartella dell'INI** (gli input `FILE.*` sono verificati per
  basename nella CWD; gli output `FILE.OUT.*` sono esclusi dal controllo).
- Dipendenze: `pandas` + `openpyxl` (xlsx) + `xlwt` (scrittura xls) + `xlrd` (rilettura xls).
  L'output `.xls` è scritto diretto con `xlwt` (pandas 3.x non scrive più il BIFF);
  estensioni maiuscole (`.XLS`, `.XLSX`) accettate.
  Driver DB per `TYPE` (verifica preventiva con messaggio di installazione):
  `sqlite3` stdlib, `pyodbc`, `oracledb`, `psycopg`, `pymysql`.
- Specifiche: `Prompt_ntj_fact.md` (file, vincolante), `prompt_ntj_fact_db.md` (database).
  Problemi/limiti noti: `ntj_fact_problems.md`.
