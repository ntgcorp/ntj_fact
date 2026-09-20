# ntj_fact

[English](#english) | [Italiano](#italiano)

<a id="english"></a>
## English

### Purpose
`ntj_fact.py` is a **ntJobsApp** that manipulates known data files in batch mode.
For a full explanation of what a ntJobsApp is (INI in, `.end` INI out, exit codes,
`acJobsApp` orchestration), see **https://github.com/ntgcorp/ntJobsApp**.

Scope of `ntj_fact`: manipulate data files automatically — **not** import/export them.
File types in scope: XLS/XLSX (Excel), CSV, JSON, XML.
First (current) version implements **XLS/XLSX only**.

### Commands

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
If the sheet already exists in the destination it is emptied first, then
source content (values and formulas, merged cells, row/column sizes) is pasted.
- `FILE.IN.XLS`: base file (copied to output).
- `FILE.IN.SHEET`: file the source sheet is read from (supports `$VARIABLES` from `[CONFIG]`).
- `SHEET.IN`: source sheet name and destination sheet name.
- `FILE.OUT.XLS`: output file (`.xls` or `.xlsx`, any case).
- Returns copied-row count + output file.

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
FILE.OUT.XLS=UTENTI_MACH0_CMLTV.XLSX
```

- Result: `<job>.end` INI file next to the input INI; exit code `0` ok,
  `1` startup/INI error, `2` one or more jobs failed.
- Run with **CWD = folder of the INI** (`FILE.*` inputs are checked by basename in CWD;
  `FILE.OUT.*` outputs are skipped by the check).
- Dependencies: `pandas` + `openpyxl` (xlsx) + `xlwt` (xls writing) + `xlrd` (xls re-read).
  `.xls` output is written directly with `xlwt` (pandas 3.x no longer writes BIFF);
  uppercase extensions (`.XLS`, `.XLSX`) are accepted.
- Spec: `Prompt_ntj_fact.md` (binding). Known issues/limits: `ntj_fact_problems.md`.

---

<a id="italiano"></a>
## Italiano

### Scopo
`ntj_fact.py` è una **ntJobsApp** che manipola in modalità batch file di dati "conosciuti".
Per la spiegazione di cosa fa una ntJobsApp (INI in ingresso, INI `.end` in uscita,
exit code, orchestrazione `acJobsApp`), vedi **https://github.com/ntgcorp/ntJobsApp**.

Scopo di `ntj_fact`: manipolare automaticamente i file di dati — **non** importarli/esportarli.
Formati previsti: XLS/XLSX (Excel), CSV, JSON, XML.
La prima versione (attuale) implementa **solo XLS/XLSX**.

### Comandi

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
Se il foglio esiste già in destinazione viene prima svuotato, poi il contenuto
sorgente (valori e formule, celle unite, dimensioni righe/colonne) viene incollato.
- `FILE.IN.XLS`: file base (copiato in output).
- `FILE.IN.SHEET`: file da cui si legge il foglio sorgente (supporta `$VARIABILI` di `[CONFIG]`).
- `SHEET.IN`: nome del foglio sorgente e di destinazione.
- `FILE.OUT.XLS`: file di output (`.xls` o `.xlsx`, maiuscolo o minuscolo).
- Ritorna il numero di righe copiate + file di output.

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
FILE.OUT.XLS=UTENTI_MACH0_CMLTV.XLSX
```

- Risultato: file INI `<job>.end` accanto all'INI di ingresso; exit code `0` tutto ok,
  `1` errore di avvio/INI, `2` uno o più job in errore.
- Lanciare con **CWD = cartella dell'INI** (gli input `FILE.*` sono verificati per
  basename nella CWD; gli output `FILE.OUT.*` sono esclusi dal controllo).
- Dipendenze: `pandas` + `openpyxl` (xlsx) + `xlwt` (scrittura xls) + `xlrd` (rilettura xls).
  L'output `.xls` è scritto diretto con `xlwt` (pandas 3.x non scrive più il BIFF);
  estensioni maiuscole (`.XLS`, `.XLSX`) accettate.
- Specifica: `Prompt_ntj_fact.md` (vincolante). Problemi/limiti noti: `ntj_fact_problems.md`.
