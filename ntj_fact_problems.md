# ntj_fact — problemi, incongruenze e bug corretti

Stato al 2026-09-19. Prima versione funzionante e testata:
`mach0/mach0_utenti_test.ini` → 255 righe (157 MATIV + 98 CMLTV), exit 0.
Comando `XLS.SHEET.APPEND` implementato e testato (xlsx/xls, solo valori, merge,
overwrite, ramo errore foglio mancante), exit 0/2 come da standard.

## Bug trovati e corretti

1. **`acJobsApp.Start` bloccava i file di output** (`acJobsApp.py`).
   Il controllo "FILE.* deve esistere" scattava anche su `FILE.OUT.XLS`,
   che per definizione non esiste prima dell'esecuzione
   (`File richiesto non presente MACH0_UTENTI.XLS`, exit 2 senza eseguire nulla).
   Fix: il controllo salta le chiavi `FILE.OUT.*` (solo input `FILE.*`).
2. **Match foglio case-sensitive** (`ntj_fact.py: sheet_resolve`).
   L'INI richiede `SHEET.OUT=UTENTI` ma i fogli reali si chiamano `utenti`.
   Con match esatto pandas non trovava il foglio.
   Fix: risoluzione case-insensitive, in scrittura si usa il nome chiesto (`UTENTI`).
3. **Estensione output case-sensitive** (`ntj_fact.py: arg_verify`).
   `MACH0_UTENTI.XLS` (maiuscolo) falliva il controllo `endswith(('.xls','.xlsx'))`.
   Fix: confronto su `lower()`.
4. **pandas 3.x non scrive piu' il formato `.xls`** (`ntj_fact.py: exec/exec_xls`).
   `df.to_excel("...XLS")` falliva con `No engine for filetype: 'XLS'`
   (supporto xlwt rimosso da pandas).
   Fix: output `.xls` scritto diretto con `xlwt`; `.xlsx` resta su pandas/openpyxl.
5. **pandas non accetta estensioni maiuscole in scrittura** (`ntj_fact.py: exec`).
   Con `FILE.OUT.XLS=MACH0_UTENTI.XLSX` falliva con
   `No engine for filetype: 'XLSX'` (match case-sensitive dentro pandas,
   NON dipendenza mancante malgrado il messaggio).
   Fix: scrittura su percorso con estensione minuscola + `os.replace`
   al nome richiesto. Stessa causa valeva per `.XLS` (coperto da xlwt diretto).
6. **Controllo preventivo dipendenze** (`ntj_fact.py: engines_verify`,
   in sequenza dopo `arg_verify`): `.xlsx` richiede `openpyxl`, `.xls`
   richiede `xlwt`. Se mancano, errore esplicito
   (`dipendenza ... non installata ... installare con pip install ...`)
   prima di leggere i file.
    Dipendenze risultanti: `pandas`, `openpyxl` (xlsx), `xlwt` (scrittura xls),
    `xlrd` (solo per rileggere/verificare xls).
7. **Tabelle query orfane → Excel "contenuto illeggibile"** (`ntj_fact.py: tables_normalize`,
    chiamata in `xls_sheet_create` dopo `load_workbook`).
    Le basi CMLTV contengono tabelle `tableType="queryTable"` (`Abilitazioni`)
    legate a `xl/queryTables/`, `xl/connections.xml` e `xl/tables/_rels/`,
    che openpyxl 3.1.5 non salva; restava il ListObject orfano e Excel
    apriva il risultato in ripristino. Fix: declassate a `worksheet`
    con `queryTableFieldId` rimosso (l'output e' uno snapshot statico,
    il legame query e' perso comunque). Verificato su copie locali
    (2 tabelle normalizzate, 361 righe, package integro).
8. **APPEND incolla solo valori, mai formule** (`ntj_fact.py: xls_sheet_append`,
    `copy_openpyxl`, `write_xlwt`; decisione D aggiornata).
    Sorgenti `.xlsx` letti con `data_only=True` (valori cached di Excel),
    output `.xls` senza piu' conversione `=...` → `xlwt.Formula`.
    Richiede workbook calcolati (senza cache la cella risulta vuota).
10. **Errori di copia non piu' silenziati** (`ntj_fact.py: copy_openpyxl`,
    `grid_from_openpyxl`, 2026-09-30).
    Ritornavano 0/righe-vuote inghiottendo l'eccezione: job in `S` con
    foglio vuoto e nessun segnale. Ora ritornano `(dati, errore)` e
    l'errore risale a video, log ed `.end` (exit 2).
9. **APPEND con SHEET.OUT diverso da SHEET.IN** (`ntj_fact.py: xls_append`,
    richiesta utente 2026-09-30).
    Nuovo parametro opzionale `SHEET.OUT` (default `SHEET.IN`): lo sheet
    sorgente si copia nel foglio di nome `SHEET.OUT`; se esiste gia'
    viene svuotato e riempito con incolla-valori (mai append).
    Variabile `out_sheet_name`, validazione titolo anche su OUT.

## Incongruenze del Prompt non correggibili nel codice (risolte per interpretazione)

5. `Prompt_ntj_fact.md` dice che `FILE.IN.SHEETS.TXT` "contiene, uno per riga,
   il nome degli sheet da leggere": in realta' contiene i **percorsi dei file**
   xlsx (uno per riga), come confermato dai txt reali e da `xls_merge.py`.
   Il codice li tratta come percorsi file.
6. La sequenza di esecuzione elencata (`arg_verify, header_read, header_check,
   read, exec`) **omette `read_list`**, senza la quale `in_files_list` resterebbe
   vuota. Il codice esegue `read_list` dopo `arg_verify`.
7. Il messaggio d'errore "foglio non trovato" cita variabili incoerenti
   (`in_sheet`, `xls_in_file_current`); il codice usa `in_sheet_name` e
   `in_file_current`: `Non trovato <sheet> nel file <file>`.
8. `xls_merge_read`: "se la prima colonna della riga corrente e' vuota esci"
   e' ambiguo (quale riga?). Implementato come in `xls_merge.py`: errore solo se
   la prima colonna e' **interamente** vuota.
9. `xls_merge_exec`: il Prompt chiede di "scrivere come ritorno numero_righe_scritte":
   implementato come `RETURN.VALUE=<n>` + `RETURN.TYPE=S` + `RETURN.FILE.XLS=<output>`
   via `jData.Return()` (standard ntJobsApp).

## Limitazioni note (non corrette, da decidere)

10. `acJobsApp.Start` verifica i `FILE.*` di input sul solo **basename nella CWD**,
    ignorando eventuali sottocartelle nel valore. L'INI mach0 funziona perche'
    si lancia dalla cartella `mach0/`. Non toccato: cambiarlo potrebbe alterare
    altre ntJobsApp.
11. Il log (`acLog.Start`) va di default nella cartella dello **script**, non in
    quella dell'INI: lanciando da `mach0/`, `mach0_utenti_test.log` nasce accanto
    a `ntj_fact.py`. Comportamento ereditato da `acJobsApp`, lasciato invariato.
12. Limiti formato `.xls` (BIFF): max 65536 righe / 256 colonne, nomi foglio a 31
    caratteri (troncato). Oltre questi limiti `exec_xls` ritorna errore esplicito.
    Per dataset grandi usare `.XLSX` in `FILE.OUT.XLS`.
13. Formattazione Excel non preservata nel merge (valori soli, header come testo);
    i NaN diventano celle vuote, i float interi (es. 138176.0) sono scritti come int.
14. I percorsi `K:\...` nei txt mach0 esistono solo nell'ambiente di produzione
    (share + `K:\Tools\pyn.cmd`); in sviluppo usare liste con percorsi locali.

## Analisi di progettazione XLS.SHEET.APPEND (2026-09-19, DA DECIDERE prima di implementare)

A. **Foglio destinazione senza nome.** La specifica elenca 3 parametri
    (`FILE.IN.XLS`, `FILE.OUT.XLS`, `SHEET.IN`) e nessun `SHEET.OUT`: il foglio
    destinazione puo' quindi chiamarsi solo come `SHEET.IN`. Scritto cosi' nel
    Prompt. Alternativa scartata salvo richiesta: parametro `SHEET.OUT` opzionale.
B. **Contraddizione crea-da-zero vs svuota-se-esiste (BLOCCANTE).**
    Se `FILE.OUT.XLS` viene sempre creato nuovo, lo sheet non puo' mai
    pre-esistere e il ramo "svuota se esiste" e' codice morto. Peggio: ogni run
    distrugge il risultato dei run precedenti, quindi accodare N sheet di N file
    in un unico OUT con N job e' impossibile.
    Proposta: `xls_out_create` crea OUT solo se manca, altrimenti apre l'esistente
    (allora "svuota se esiste" ha senso). Da confermare.
C. **"Svuotato" vs "cancella quello gia' esistente" (BLOCCANTE).**
    La descrizione del comando dice "svuotato", `xls_sheet_create` dice "cancella".
    Non sono equivalenti: svuotare (clear contenuto) conserva posizione del foglio,
    colore tab, freeze panes, impostazioni stampa e filtri; cancellare+ricreare
    azzera tutto ma garantisce pulizia totale (anche merge/stili orfani).
    Proposta: svuota (come dice la frase principale del comando). Da confermare.
D. **Perimetro della copia non definito.**
    "Formule comprese" = valori + formule come testo (openpyxl le copia cosi';
    i riferimenti relativi restano validi se il foglio keepa lo stesso nome;
    Excel ricalcola all'apertura). Non specificati: stili, celle unite,
    larghezze colonne/altezze righe, freeze, filtri, immagini/grafici
    (questi ultimi openpyxl non li copia in modo banale).
    Proposta: valori + formule + celle unite + dimensioni colonne/righe;
    niente stili/immagini/grafici. Da confermare.
E. **Input `.xls` (BIFF) e "formule comprese" incompatibili.**
    openpyxl non legge `.xls`; xlrd legge solo i valori calcolati, NON le formule.
    Proposta: APPEND preserva le formule solo da `.xlsx`; da `.xls` copia i valori
    (documentato) — alternativa: errore esplicito su input `.xls`. Da confermare.
F. **Ritorno non specificato.** Proposta (come MERGE): `RETURN.VALUE` = numero
    righe copiate + `RETURN.FILE.XLS` = output, via `jData.Return()`.
G. Frase "inserisce un file xls o xls lo sheet di un altro file" incomprensibile:
    interpretata come "copia SHEET.IN da FILE.IN.XLS in FILE.OUT.XLS" (scritto cosi').
H. Refuso "SHHET.IN" → `SHEET.IN` (scritto corretto nel Prompt).
I. Con openpyxl `max_row`/`max_column` possono essere gonfiati da formattazione
    residua: in implementazione delimitare l'area copiata all'used-range reale,
    altrimenti righe/colonne vuote fantasma finiscono nell'output.

## Decisioni XLS.SHEET.APPEND (2026-09-19, confermate dall'utente)

- B: FILE.OUT.XLS sempre sovrascritto (niente append multipli nello stesso OUT).
- C: sheet esistente svuotato (non cancellato/ricreato).
- D: copia solo valori (mai formule: incolla-valori) + celle unite + dimensioni righe/colonne (modifica 2026-09-30, richiesta utente).
- E: da input `.xls` si copiano i valori (limite documentato).
- F: ritorno `RETURN.VALUE` = righe copiate + `RETURN.FILE.XLS` (come MERGE).

## Modifica specifica APPEND a due file (2026-09-19, richiesta utente)
- `FILE.OUT.XLS` nasce come copia binaria immediata di `FILE.IN.XLS`
  (base); lo sheet `SHEET.IN` si legge dal diverso file `FILE.IN.SHEET`.
  Vecchia semantica "OUT creato vuoto" superata (era anche contraddittoria
  con "svuota se esiste": ora lo sheet puo' pre-esistere via base).
- Restano validi: OUT sempre rigenerato ogni run, sheet esistente svuotato,
  perimetro copia D, valori da `.xls`.
- ATTENZIONE uso reale: se la base contiene gia' uno sheet di nome `SHEET.IN`,
  viene SVUOTATO e sostituito dal contenuto di `FILE.IN.SHEET`.
  E' il caso di `mach0_ya54m0_test.ini`: la base `MACH0_UTENTI.XLSX` (da MERGE,
  sheet `UTENTI`) + `SHEET.IN=UTENTI` da `YA54M0.XLSX` → lo sheet `UTENTI`
  della base viene sostituito dai dati YA54M0, non affiancato.
- Con output `.xls` la base viene riletta via xlrd e riscritta via xlwt:
  formule e formattazione degli sheet della base degradano a valori
  (limite BIFF, come punto E).

## Fix espansione prima del controllo file (2026-09-20)

- **`acJobsApp.Start` validava i `FILE.*` prima di espandere le `$variabili`.**
  Con `FILE.IN.SHEET=$YA54M0` (variabile in CONFIG) abortiva con
  `File richiesto non presente $YA54M0` (exit 2) senza mai espandere,
  anche con `YA54M0.XLSX` presente nella cartella di lancio.
  Fix: espansione CONFIG + sezioni spostata PRIMA della verifica `FILE.*`,
  che ora gira sui valori gia' espansi. Testato con riproduzione dedicata.
- Resta il limite basename: dopo l'espansione, di un percorso assoluto
  (es. `K:\_Statistiche\YA54M0.XLSX`) si verifica il solo basename nella CWD.

## Comandi DB: rename, STRUCT facoltativa, DB.EXPORT (2026-09-30)

- `IMPORT`→`DB.IMPORT`; `STRUCT`→`FILE.IN.STRUCT`, `SQL.FILE`→`FILE.IN.SQL`
  (regola: input `FILE.IN.*`, output `FILE.OUT.*`; ora anche verificati da
  `acJobsApp.Start` in CWD).
- `FILE.IN.STRUCT` facoltativa in `DB.IMPORT`: senza, `TABLE` obbligatorio
  e colonne da introspezione (`SELECT WHERE 1=0`, tutte nullable);
  colonne `auto` mai richieste nel file.
- Nuovo `DB.EXPORT` (`OUT.QUERY` SELECT + `FILE.OUT.QUERY` + `OUT.TYPE`
  CSV|JSON con match estensione).
- `jData.Return` in tutti i comandi DB (`RETURN.VALUE` + `RETURN.FILE.*`);
  `dictDb` popolato da sezioni `[DB_*]`; dispatcher esplicito in
  `ntj_fact.py` (fix: comandi vuoti/sconosciuti passavano in silenzio
  con modulo DB caricato). Testato end-to-end su SQLITE.