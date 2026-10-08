"""
ntj_fact.py - ntJobsApp per la manipolazione di file di dati conosciuti.

Prima versione: solo XLS/XLSX (comando XLS.TAB.MERGE).
Specifiche: Prompt_ntj_fact.md (file) e prompt_ntj_fact_db.md (database).
Contesto ntJobsApp e convenzioni: Prompt_ntjobsapp_light.md.

Comandi file: XLS.TAB.MERGE, XLS.SHEET.APPEND.
Comandi database (se ntj_fact_db.py importato): DB.OPEN, DB.CREATE, DB.CLOSE,
SQL.EXEC, TAB.ZAP, TAB.CREATE.JSON, IMPORT. Vedi prompt_ntj_fact_db.md.

Uso:
    python ntj_fact.py <file_job.ini>
"""

import os
import sys
from typing import Dict, List, Any, Optional

try:
    import pandas as pd
except ImportError:
    print("ntj_fact: Errore dipendenza pandas non installata "
          "(serve anche openpyxl per xlsx, xlwt per xls)")
    sys.exit(1)

from acJobsApp import acJobsApp, ErrorProc, Timestamp, NormalizePath

NTJ_FACT_VER = "20260919"

jData = acJobsApp()

# Flag modulo DB: True se ntj_fact_db.py e' stato importato con successo
_db_module_available = False
try:
    import ntj_fact_db
    _db_module_available = True
except ImportError:
    pass


class xls_merge:
    """
    Unisce file xls/xlsx con identico header in un unico file di risultato.

    Variabili di istanza (da Prompt_ntj_fact.md):
    - in_files_name: file txt con un percorso xlsx per riga
    - in_sheet_name: nome del foglio da leggere/scrivere
    - out_xls_name: file Excel di output
    - in_files_list: array dei percorsi letti dal txt
    - dictXlsHeaders: {percorso: [header]} (prima riga di ogni file)
    - dictXlsDati: {percorso: [[righe]]} (contenuto dati di ogni file)
    """

    def __init__(self) -> None:
        """Inizializza le variabili della classe."""
        sProc = "__init__"
        print("Sto eseguendo " + sProc)
        self.in_files_name = ""
        self.in_sheet_name = ""
        self.out_xls_name = ""
        self.in_files_list: List[str] = []
        self.dictXlsHeaders: Dict[str, List[Any]] = {}
        self.dictXlsDati: Dict[str, List[List[Any]]] = {}
        self.num_righe_lette = 0
        self.num_righe_scritte = 0

    def arg_verify(self, dictJob: Dict[str, str]) -> str:
        """
        Legge e verifica i parametri del job.

        Si aspetta FILE.IN.SHEETS.TXT, SHEET.OUT, FILE.OUT.XLS.
        """
        sProc = "arg_verify"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            self.in_files_name = str(dictJob.get("FILE.IN.SHEETS.TXT", "")).strip()
            self.in_sheet_name = str(dictJob.get("SHEET.OUT", "")).strip()
            self.out_xls_name = str(dictJob.get("FILE.OUT.XLS", "")).strip()
            if not self.in_files_name:
                sResult = "Parametro FILE.IN.SHEETS.TXT non precisato"
            elif not os.path.isfile(NormalizePath(self.in_files_name)):
                sResult = "Il file '" + self.in_files_name + "' non esiste."
            elif not self.in_sheet_name:
                sResult = "Parametro SHEET.OUT non precisato"
            elif not self.out_xls_name:
                sResult = "Parametro FILE.OUT.XLS non precisato"
            elif not self.out_xls_name.lower().endswith((".xls", ".xlsx")):
                sResult = ("Il file di output '" + self.out_xls_name
                           + "' deve avere estensione .xls o .xlsx.")
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def engines_verify(self) -> str:
        """
        Verifica che la libreria di scrittura per l'output sia installata.

        .xlsx richiede openpyxl, .xls richiede xlwt. Fallisce prima di
        leggere i file, con messaggio che indica cosa installare.
        """
        sProc = "engines_verify"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            if self.out_xls_name.lower().endswith(".xlsx"):
                try:
                    import openpyxl  # noqa: F401
                except ImportError:
                    sResult = ("dipendenza openpyxl non installata: serve per "
                               "scrivere .xlsx (installare con pip install openpyxl)")
            elif self.out_xls_name.lower().endswith(".xls"):
                try:
                    import xlwt  # noqa: F401
                except ImportError:
                    sResult = ("dipendenza xlwt non installata: serve per "
                               "scrivere .xls (installare con pip install xlwt)")
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def read_list(self) -> str:
        """
        Legge il txt in in_files_name, una riga per percorso.

        Salva ogni riga non vuota in in_files_list.
        """
        sProc = "read_list"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            self.in_files_list = []
            with open(NormalizePath(self.in_files_name), "r", encoding="utf-8-sig") as f:
                for sLine in f:
                    sLine = sLine.strip()
                    if sLine:
                        self.in_files_list.append(sLine)
            if not self.in_files_list:
                sResult = "Nessun file elencato in " + self.in_files_name
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def sheet_resolve(self, xls_file: Any) -> str:
        """
        Restituisce il nome reale del foglio (match case-insensitive).

        Restituisce "" se il foglio non esiste nel workbook.
        """
        sProc = "sheet_resolve"
        print("Sto eseguendo " + sProc)
        try:
            sWant = str(self.in_sheet_name).lower()
            for sName in xls_file.sheet_names:
                if str(sName).lower() == sWant:
                    return str(sName)
            return ""
        except Exception:
            return ""

    def header_read(self) -> str:
        """
        Legge la prima riga di ogni file in dictXlsHeaders.

        Chiave: percorso corrente; valore: array delle colonne lette.
        """
        sProc = "header_read"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            self.dictXlsHeaders = {}
            for in_file_current in self.in_files_list:
                try:
                    xls_file = pd.ExcelFile(NormalizePath(in_file_current))
                except Exception as e:
                    sResult = ("Errore durante la lettura del file "
                               + in_file_current + ": " + str(e))
                    break
                sSheetReal = self.sheet_resolve(xls_file)
                if not sSheetReal:
                    sResult = ("Non trovato " + self.in_sheet_name
                               + " nel file " + in_file_current)
                    break
                try:
                    df = pd.read_excel(xls_file, sheet_name=sSheetReal,
                                       header=None, nrows=1)
                except Exception as e:
                    sResult = ("Errore durante la lettura del file "
                               + in_file_current + ": " + str(e))
                    break
                header_row = df.values.tolist()[0]
                self.dictXlsHeaders[in_file_current] = header_row
                print("Header letto da " + in_file_current + ": "
                      + str(header_row))
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def header_check(self) -> str:
        """
        Verifica che gli header di tutti i file siano uguali.

        Confronta posizione per posizione con l'header del primo file.
        """
        sProc = "header_check"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            keys = list(self.dictXlsHeaders.keys())
            if not keys:
                sResult = "Nessun file Excel trovato nel dizionario."
                return ErrorProc(sResult, sProc)
            xls_header_base = self.dictXlsHeaders[keys[0]]
            print("Header di riferimento (da " + keys[0] + "): "
                  + str(xls_header_base))
            for key in keys[1:]:
                current_header = self.dictXlsHeaders[key]
                print("Confronto con header di " + key + ": "
                      + str(current_header))
                if len(current_header) != len(xls_header_base):
                    sResult = ("Header del file " + key + " non corretto, "
                               "numero di colonne differente: "
                               + str(len(current_header))
                               + " (atteso: " + str(len(xls_header_base)) + ")")
                    break
                for i, (base_col, current_col) in enumerate(
                        zip(xls_header_base, current_header)):
                    if base_col != current_col:
                        sResult = ("Header del file " + key + " non corretto, "
                                   "nome campo differente: '" + str(current_col)
                                   + "' (atteso: '" + str(base_col)
                                   + "') alla posizione " + str(i + 1))
                        break
                if sResult:
                    break
            if not sResult:
                print("Tutti gli header sono uguali!")
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def read(self) -> str:
        """
        Legge il contenuto dati di ogni file in dictXlsDati.

        Stesse chiavi di dictXlsHeaders; errore se la prima colonna
        di un file e' interamente vuota.
        """
        sProc = "read"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            self.dictXlsDati = {}
            self.num_righe_lette = 0
            for in_file_current in self.dictXlsHeaders.keys():
                try:
                    xls_file = pd.ExcelFile(NormalizePath(in_file_current))
                    sSheetReal = self.sheet_resolve(xls_file)
                    df = pd.read_excel(xls_file, sheet_name=sSheetReal)
                except Exception as e:
                    sResult = ("Errore durante la lettura del file "
                               + in_file_current + ": " + str(e))
                    break
                if df.iloc[:, 0].isnull().all():
                    sResult = ("La prima colonna del file " + in_file_current
                               + " e' vuota.")
                    break
                rows = df.values.tolist()
                self.dictXlsDati[in_file_current] = rows
                self.num_righe_lette += len(rows)
                print("Lette righe dal file excel: " + str(len(rows))
                      + " (" + in_file_current + ")")
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def exec(self) -> str:
        """
        Scrive tutti i dati in un unico Excel con un solo header.

        Unico foglio di nome in_sheet_name nel file out_xls_name.
        L'estensione .xls usa xlwt diretto (pandas non scrive piu'
        il formato BIFF); .xlsx usa pandas/openpyxl.
        """
        sProc = "exec"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            keys = list(self.dictXlsDati.keys())
            if not keys:
                sResult = "Nessun dato da scrivere."
                return ErrorProc(sResult, sProc)
            all_data: List[List[Any]] = []
            for data in self.dictXlsDati.values():
                all_data.extend(data)
            header = self.dictXlsHeaders[keys[0]]
            sOut = NormalizePath(self.out_xls_name)
            sDir = os.path.dirname(sOut)
            if sDir:
                os.makedirs(sDir, exist_ok=True)
            if sOut.lower().endswith(".xls") and not sOut.lower().endswith(".xlsx"):
                sResult = self.exec_xls(sOut, header, all_data)
                if sResult:
                    return ErrorProc(sResult, sProc)
            else:
                # pandas confronta l'estensione in modo case-sensitive:
                # si scrive su percorso con estensione minuscola e poi si
                # rinomina al nome richiesto (es. MACH0_UTENTI.XLSX).
                sBase, sExt = os.path.splitext(sOut)
                sOutWrite = sBase + sExt.lower()
                df = pd.DataFrame(all_data, columns=header)
                df.to_excel(sOutWrite, sheet_name=self.in_sheet_name, index=False)
                if os.path.abspath(sOutWrite) != os.path.abspath(sOut):
                    os.replace(sOutWrite, sOut)
            self.num_righe_scritte = len(all_data)
            print("Lette righe dal file excel: " + str(self.num_righe_scritte))
            print("File '" + self.out_xls_name + "' creato con successo.")
        except Exception as e:
            sResult = "Errore durante il salvataggio del file Excel: " + str(e)
        return ErrorProc(sResult, sProc)

    def exec_xls(self, sOut: str, header: List[Any], rows: List[List[Any]]) -> str:
        """
        Scrive header + righe in formato .xls (BIFF) tramite xlwt.

        NaN/None diventano celle vuote; i limiti del formato .xls
        (65536 righe, 256 colonne) sono verificati prima di scrivere.
        """
        sProc = "exec_xls"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            import xlwt
        except ImportError:
            return ErrorProc("dipendenza xlwt non installata", sProc)
        try:
            if len(header) > 256:
                sResult = "Troppe colonne per il formato .xls: " + str(len(header))
                return ErrorProc(sResult, sProc)
            if len(rows) + 1 > 65536:
                sResult = "Troppe righe per il formato .xls: " + str(len(rows))
                return ErrorProc(sResult, sProc)
            wb = xlwt.Workbook()
            ws = wb.add_sheet(self.in_sheet_name[:31])
            for j, col in enumerate(header):
                ws.write(0, j, "" if col is None else str(col))
            for i, row in enumerate(rows, start=1):
                for j in range(len(header)):
                    v: Any = row[j] if j < len(row) else None
                    if v is None:
                        continue
                    try:
                        if isinstance(v, float) and pd.isna(v):
                            continue
                    except Exception:
                        pass
                    if isinstance(v, float) and v.is_integer():
                        v = int(v)
                    ws.write(i, j, v)
            wb.save(sOut)
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)


class xls_append:
    """
    Copia lo sheet SHEET.IN da FILE.IN.XLS in FILE.OUT.XLS (foglio omonimo).

    Variabili di istanza:
    - in_xls_name: file xls/xlsx base (deve esistere, copiato in output)
    - in_sheet_file: diverso file xls/xlsx da cui si legge SHEET.IN
    - out_xls_name: file xls/xlsx destinazione (copia della base + sheet)
    - in_sheet_name: nome dello sheet sorgente (in FILE.IN.SHEET)
    - out_sheet_name: nome del foglio destinazione (default in_sheet_name)
    - num_righe_copiate: righe non vuote copiate
    - wb_out/ws_out: workbook/sheet di output (solo flusso xlsx)
    """

    def __init__(self) -> None:
        """Inizializza le variabili della classe."""
        sProc = "__init__"
        print("Sto eseguendo " + sProc)
        self.in_xls_name = ""
        self.in_sheet_file = ""
        self.out_xls_name = ""
        self.in_sheet_name = ""
        self.out_sheet_name = ""
        self.num_righe_copiate = 0
        self.wb_out: Any = None
        self.ws_out: Any = None

    def is_xlsx(self, sFile: str) -> bool:
        """True se il percorso termina con .xlsx (qualsiasi case)."""
        sProc = "is_xlsx"
        print("Sto eseguendo " + sProc)
        return sFile.lower().endswith(".xlsx")

    def is_xls(self, sFile: str) -> bool:
        """True se il percorso termina con .xls ma non .xlsx (qualsiasi case)."""
        sProc = "is_xls"
        print("Sto eseguendo " + sProc)
        sLow = sFile.lower()
        return sLow.endswith(".xls") and not sLow.endswith(".xlsx")

    def sheet_title_verify(self, sName: str = "") -> str:
        """
        Verifica che sName sia un nome foglio Excel valido.

        Senza sName verifica in_sheet_name. Max 31 caratteri, senza []:*?/\\.
        """
        sProc = "sheet_title_verify"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            if not sName:
                sName = self.in_sheet_name
            if len(sName) > 31:
                sResult = ("Nome sheet troppo lungo (max 31 caratteri): "
                           + sName)
            elif any(c in sName for c in "[]:*?/\\"):
                sResult = ("Nome sheet con caratteri non validi ([]:*?/\\): "
                           + sName)
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def xls_append_args(self, dictJob: Dict[str, str]) -> str:
        """
        Legge e verifica i parametri del job.

        Si aspetta FILE.IN.XLS (base esistente), FILE.IN.SHEET (esistente),
        SHEET.IN, SHEET.OUT (opzionale, default SHEET.IN),
        FILE.OUT.XLS (estensione .xls/.xlsx, qualsiasi case).
        Verifica anche le librerie richieste dalle estensioni in gioco.
        """
        sProc = "xls_append_args"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            self.in_xls_name = str(dictJob.get("FILE.IN.XLS", "")).strip()
            self.in_sheet_file = str(dictJob.get("FILE.IN.SHEET", "")).strip()
            self.in_sheet_name = str(dictJob.get("SHEET.IN", "")).strip()
            self.out_sheet_name = str(dictJob.get("SHEET.OUT", "")).strip()
            self.out_xls_name = str(dictJob.get("FILE.OUT.XLS", "")).strip()
            if not self.in_xls_name:
                sResult = "Parametro FILE.IN.XLS non precisato"
            elif not os.path.isfile(NormalizePath(self.in_xls_name)):
                sResult = "Il file '" + self.in_xls_name + "' non esiste."
            elif not self.in_sheet_file:
                sResult = "Parametro FILE.IN.SHEET non precisato"
            elif not os.path.isfile(NormalizePath(self.in_sheet_file)):
                sResult = "Il file '" + self.in_sheet_file + "' non esiste."
            elif not self.in_sheet_name:
                sResult = "Parametro SHEET.IN non precisato"
            elif not self.out_xls_name:
                sResult = "Parametro FILE.OUT.XLS non precisato"
            elif not (self.is_xlsx(self.out_xls_name)
                      or self.is_xls(self.out_xls_name)):
                sResult = ("Il file di output '" + self.out_xls_name
                           + "' deve avere estensione .xls o .xlsx.")
            if not sResult:
                if not self.out_sheet_name:
                    self.out_sheet_name = self.in_sheet_name
                sResult = self.sheet_title_verify()
            if not sResult:
                sResult = self.sheet_title_verify(self.out_sheet_name)
            if not sResult:
                sResult = self.append_engines_verify()
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def append_engines_verify(self) -> str:
        """
        Verifica le librerie richieste dalle estensioni in gioco.

        Lettura/scrittura .xlsx: openpyxl. Lettura .xls (base, sheet
        sorgente o rilettura output): xlrd. Scrittura .xls: xlwt.
        """
        sProc = "append_engines_verify"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            bXlsx = (self.is_xlsx(self.in_xls_name)
                     or self.is_xlsx(self.in_sheet_file)
                     or self.is_xlsx(self.out_xls_name))
            bXlsIn = (self.is_xls(self.in_xls_name)
                      or self.is_xls(self.in_sheet_file))
            if bXlsx:
                try:
                    import openpyxl  # noqa: F401
                except ImportError:
                    sResult = ("dipendenza openpyxl non installata: serve per "
                               "leggere/scrivere .xlsx "
                               "(installare con pip install openpyxl)")
                    return ErrorProc(sResult, sProc)
            if bXlsIn or self.is_xls(self.out_xls_name):
                try:
                    import xlrd  # noqa: F401
                except ImportError:
                    sResult = ("dipendenza xlrd non installata: serve per "
                               "leggere .xls (installare con pip install xlrd)")
                    return ErrorProc(sResult, sProc)
            if self.is_xls(self.out_xls_name):
                try:
                    import xlwt  # noqa: F401
                except ImportError:
                    sResult = ("dipendenza xlwt non installata: serve per "
                               "scrivere .xls (installare con pip install xlwt)")
                    return ErrorProc(sResult, sProc)
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def xls_out_create(self) -> str:
        """
        Crea il file di output come copia di FILE.IN.XLS.

        Copia binaria (qualsiasi formato), eventualmente sovrascrivendo
        il precedente. Lo sheet viene inserito in xls_sheet_append.
        """
        sProc = "xls_out_create"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            import shutil
            sIn = NormalizePath(self.in_xls_name)
            sOut = NormalizePath(self.out_xls_name)
            sDir = os.path.dirname(sOut)
            if sDir:
                os.makedirs(sDir, exist_ok=True)
            self.wb_out = None
            self.ws_out = None
            shutil.copyfile(sIn, sOut)
            print("File di output '" + self.out_xls_name + "' creato da '"
                  + self.in_xls_name + "'.")
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def sheet_resolve_out(self) -> str:
        """
        Restituisce il titolo reale del foglio destinazione (match case-insensitive).

        Restituisce "" se non esiste.
        """
        sProc = "sheet_resolve_out"
        print("Sto eseguendo " + sProc)
        try:
            sWant = str(self.out_sheet_name).lower()
            for sName in self.wb_out.sheetnames:
                if str(sName).lower() == sWant:
                    return str(sName)
            return ""
        except Exception:
            return ""

    def tables_normalize(self) -> str:
        """
        Declassa le tabelle query in tabelle standard.

        openpyxl non salva queryTables/connections: una tabella
        tableType=queryTable senza le sue parti scatena in Excel
        il ripristino "contenuto illeggibile". L'output e' uno
        snapshot statico, il legame query e' perso comunque.
        """
        sProc = "tables_normalize"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            nNorm = 0
            for ws in self.wb_out.worksheets:
                oTables = getattr(ws, "_tables", [])
                if isinstance(oTables, dict):
                    oTables = list(oTables.values())
                for oTab in list(oTables):
                    sType = str(getattr(oTab, "tableType", "") or "")
                    if sType != "queryTable":
                        continue
                    oTab.tableType = "worksheet"
                    for oCol in list(getattr(oTab, "tableColumns", []) or []):
                        if hasattr(oCol, "queryTableFieldId"):
                            oCol.queryTableFieldId = None
                    nNorm += 1
            print("Tabelle query normalizzate: " + str(nNorm))
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def sheet_is_pristine(self, ws: Any) -> bool:
        """True se il foglio e' vuoto (una cella vuota, nessun merge)."""
        sProc = "sheet_is_pristine"
        print("Sto eseguendo " + sProc)
        try:
            if len(list(ws.merged_cells.ranges)) > 0:
                return False
            if ws.max_row != 1 or ws.max_column != 1:
                return False
            return ws.cell(row=1, column=1).value is None
        except Exception:
            return False

    def sheet_clear(self, ws: Any) -> str:
        """
        Svuota il contenuto del foglio (valori + merge), mantiene il foglio.
        """
        sProc = "sheet_clear"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            for sRange in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(sRange))
            if ws.max_row >= 1:
                ws.delete_rows(1, ws.max_row)
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def xls_sheet_create(self) -> str:
        """
        Predispone lo sheet destinazione out_sheet_name nell'output.

        Apre FILE.OUT.XLS (copia della base): se lo sheet esiste viene
        svuotato, se manca viene creato. Gli altri sheet restano invariati.
        Per output .xls verifica solo i limiti del formato (il foglio
        nasce in scrittura da base + sheet sorgente).
        """
        sProc = "xls_sheet_create"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            if not self.is_xlsx(self.out_xls_name):
                return ErrorProc(sResult, sProc)
            from openpyxl import load_workbook
            self.wb_out = load_workbook(NormalizePath(self.out_xls_name))
            sResult = self.tables_normalize()
            if sResult:
                return ErrorProc(sResult, sProc)
            sFound = self.sheet_resolve_out()
            if sFound:
                self.ws_out = self.wb_out[sFound]
                sResult = self.sheet_clear(self.ws_out)
                if sResult:
                    return ErrorProc(sResult, sProc)
                print("Sheet '" + sFound + "' svuotato.")
                return ErrorProc(sResult, sProc)
            self.ws_out = self.wb_out.create_sheet(self.out_sheet_name)
            print("Sheet '" + self.out_sheet_name + "' creato.")
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def used_range(self, ws: Any) -> tuple:
        """
        Restituisce (max_row, max_col) reali del foglio.

        Stringe max_row/max_column di openpyxl escludendo code vuote
        (possono essere gonfiate da formattazione residua).
        """
        sProc = "used_range"
        print("Sto eseguendo " + sProc)
        try:
            from openpyxl.utils import range_boundaries
            max_r = ws.max_row or 0
            max_c = ws.max_column or 0
            merged_cells = set()
            for sRange in ws.merged_cells.ranges:
                min_c, min_r, max_cc, max_rr = range_boundaries(str(sRange))
                for r in range(min_r, max_rr + 1):
                    for c in range(min_c, max_cc + 1):
                        merged_cells.add((r, c))
            while max_r > 0 and not any(
                    ws.cell(row=max_r, column=c).value is not None
                    or (max_r, c) in merged_cells
                    for c in range(1, max_c + 1)):
                max_r -= 1
            while max_c > 0 and not any(
                    ws.cell(row=r, column=max_c).value is not None
                    or (r, max_c) in merged_cells
                    for r in range(1, max_r + 1)):
                max_c -= 1
            return (max_r, max_c)
        except Exception:
            return (ws.max_row or 0, ws.max_column or 0)

    def copy_openpyxl(self, ws_in: Any, ws_out: Any) -> tuple:
        """
        Copia solo valori (mai formule: incolla-valori), merge e dimensioni.

        Il foglio sorgente va letto data_only=True dal chiamante.
        Restituisce (righe_copiate, errore): l'errore non e' mai
        silenziato, va a video e nel log via sequenza chiamante.
        """
        sProc = "copy_openpyxl"
        print("Sto eseguendo " + sProc)
        try:
            from openpyxl.utils import get_column_letter
            max_r, max_c = self.used_range(ws_in)
            nRows = 0
            for r in range(1, max_r + 1):
                bHas = False
                for c in range(1, max_c + 1):
                    v = ws_in.cell(row=r, column=c).value
                    if v is not None:
                        ws_out.cell(row=r, column=c, value=v)
                        bHas = True
                if bHas:
                    nRows += 1
            for sRange in ws_in.merged_cells.ranges:
                ws_out.merge_cells(str(sRange))
            for sLetter, dim in ws_in.column_dimensions.items():
                if dim.width is not None:
                    ws_out.column_dimensions[sLetter].width = dim.width
            for nIdx, dim in ws_in.row_dimensions.items():
                if dim.height is not None:
                    ws_out.row_dimensions[nIdx].height = dim.height
            return (nRows, "")
        except Exception as e:
            return (0, ErrorProc(str(e), sProc))

    def copy_xlrd_to_openpyxl(self, sFile: str, sSheet: str, ws_out: Any) -> tuple:
        """
        Copia valori (no formule, limite BIFF) da .xls a sheet openpyxl.

        Restituisce (righe_copiate, errore).
        """
        sProc = "copy_xlrd_to_openpyxl"
        print("Sto eseguendo " + sProc)
        try:
            import xlrd
            from openpyxl.utils import get_column_letter
            wb = xlrd.open_workbook(NormalizePath(sFile), formatting_info=False)
            sh = None
            for i in range(wb.nsheets):
                if wb.sheet_names()[i].lower() == sSheet.lower():
                    sh = wb.sheet_by_index(i)
                    break
            if sh is None:
                return (0, ErrorProc("Non trovato " + self.in_sheet_name
                                     + " nel file " + sFile, sProc))
            nRows = 0
            for r in range(sh.nrows):
                bHas = False
                for c in range(sh.ncols):
                    v = sh.cell_value(r, c)
                    if v == "":
                        continue
                    if isinstance(v, float) and v.is_integer():
                        v = int(v)
                    ws_out.cell(row=r + 1, column=c + 1, value=v)
                    bHas = True
                if bHas:
                    nRows += 1
            for (rlo, rhi, clo, chi) in sh.merged_cells:
                ws_out.merge_cells(start_row=rlo + 1, start_column=clo + 1,
                                   end_row=rhi, end_column=chi)
            return (nRows, "")
        except Exception as e:
            return (0, ErrorProc(str(e), sProc))

    def read_xls_book(self, sFile: str) -> tuple:
        """
        Legge tutti gli sheet di un .xls (titoli, valori, merge).

        Restituisce (booksheets, errore); booksheets e' una lista di
        (titolo, rows, merges, {}, {}).
        """
        sProc = "read_xls_book"
        print("Sto eseguendo " + sProc)
        try:
            import xlrd
            wb = xlrd.open_workbook(NormalizePath(sFile), formatting_info=False)
            booksheets = []
            for i in range(wb.nsheets):
                sh = wb.sheet_by_index(i)
                rows = []
                for r in range(sh.nrows):
                    row = []
                    for c in range(sh.ncols):
                        v = sh.cell_value(r, c)
                        row.append(None if v == "" else v)
                    rows.append(row)
                merges = [(rlo, clo, rhi - 1, chi - 1)
                          for (rlo, rhi, clo, chi) in sh.merged_cells]
                booksheets.append((sh.name, rows, merges, {}, {}))
            return (booksheets, "")
        except Exception as e:
            return ([], ErrorProc(str(e), sProc))

    def write_xlwt(self, sOut: str, booksheets: list) -> str:
        """
        Scrive uno o piu' sheet in formato .xls tramite xlwt.

        booksheets: lista di (titolo, rows, merges, col_widths, row_heights).
        Solo valori, mai formule (incolla-valori): le stringhe sono
        scritte come testo anche se iniziano con "=".
        """
        sProc = "write_xlwt"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            import xlwt
            from openpyxl.utils import column_index_from_string
            nRowsTot = sum(len(rows) for (_, rows, _, _, _) in booksheets)
            nColsMax = max([max([len(r) for r in rows] + [0])
                            for (_, rows, _, _, _) in booksheets] + [0])
            if nColsMax > 256:
                return ErrorProc("Troppe colonne per il formato .xls: "
                                 + str(nColsMax), sProc)
            if nRowsTot > 65536:
                return ErrorProc("Troppe righe per il formato .xls: "
                                 + str(nRowsTot), sProc)
            wb = xlwt.Workbook()
            for (sTitle, rows, merges, col_widths, row_heights) in booksheets:
                ws = wb.add_sheet(str(sTitle)[:31])
                covered = set()
                for (r1, c1, r2, c2) in merges:
                    for r in range(r1, r2 + 1):
                        for c in range(c1, c2 + 1):
                            covered.add((r, c))
                for sLetter, w in col_widths.items():
                    try:
                        ws.col(column_index_from_string(sLetter) - 1).width = int(
                            float(w) * 256)
                    except Exception:
                        pass
                for nIdx, h in row_heights.items():
                    try:
                        ws.row(nIdx - 1).height = int(float(h) * 20)
                    except Exception:
                        pass
                for i, row in enumerate(rows):
                    for j, v in enumerate(row):
                        if (i, j) in covered:
                            continue
                        if v is None or v == "":
                            continue
                        if isinstance(v, float) and v.is_integer():
                            v = int(v)
                        ws.write(i, j, v)
                for (r1, c1, r2, c2) in merges:
                    v = ""
                    if r1 < len(rows) and c1 < len(rows[r1]):
                        v = rows[r1][c1]
                        if v is None:
                            v = ""
                    ws.write_merge(r1, r2, c1, c2, v)
            wb.save(sOut)
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)

    def grid_from_openpyxl(self, ws_in: Any) -> tuple:
        """
        Estrae da sheet openpyxl (griglia, merge, dimensioni).

        Restituisce (rows, merges, col_widths, row_heights, errore):
        l'errore non e' mai silenziato, va a video e nel log
        via sequenza chiamante.
        """
        sProc = "grid_from_openpyxl"
        print("Sto eseguendo " + sProc)
        try:
            from openpyxl.utils import get_column_letter
            max_r, max_c = self.used_range(ws_in)
            rows = []
            for r in range(1, max_r + 1):
                rows.append([ws_in.cell(row=r, column=c).value
                             for c in range(1, max_c + 1)])
            merges = []
            for sRange in ws_in.merged_cells.ranges:
                from openpyxl.utils import range_boundaries
                min_c, min_r, max_cc, max_rr = range_boundaries(str(sRange))
                merges.append((min_r - 1, min_c - 1, max_rr - 1, max_cc - 1))
            col_widths = {sL: d.width for sL, d in
                          ws_in.column_dimensions.items()
                          if d.width is not None}
            row_heights = {nI: d.height for nI, d in
                           ws_in.row_dimensions.items()
                           if d.height is not None}
            return (rows, merges, col_widths, row_heights, "")
        except Exception as e:
            return ([], [], {}, {}, ErrorProc(str(e), sProc))

    def grid_from_xlrd(self, sFile: str, sSheet: str) -> tuple:
        """
        Estrae da sheet .xls (griglia valori, merge).

        Restituisce (rows, merges, col_widths, row_heights, errore).
        """
        sProc = "grid_from_xlrd"
        print("Sto eseguendo " + sProc)
        try:
            import xlrd
            wb = xlrd.open_workbook(NormalizePath(sFile), formatting_info=False)
            sh = None
            for i in range(wb.nsheets):
                if wb.sheet_names()[i].lower() == sSheet.lower():
                    sh = wb.sheet_by_index(i)
                    break
            if sh is None:
                return ([], [], {}, {},
                        ErrorProc("Non trovato " + self.in_sheet_name
                                  + " nel file " + sFile, sProc))
            rows = []
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    v = sh.cell_value(r, c)
                    row.append(None if v == "" else v)
                rows.append(row)
            merges = [(rlo, clo, rhi - 1, chi - 1)
                      for (rlo, rhi, clo, chi) in sh.merged_cells]
            return (rows, merges, {}, {}, "")
        except Exception as e:
            return ([], [], {}, {}, ErrorProc(str(e), sProc))

    def xls_sheet_append(self) -> str:
        """
        Copia SHEET.IN da FILE.IN.SHEET in FILE.OUT.XLS e salva
        (solo valori, mai formule: incolla-valori).

        Output xlsx via openpyxl (vale anche con sorgente xls, solo valori);
        output xls via xlwt riscrivendo tutti gli sheet della base.
        """
        sProc = "xls_sheet_append"
        print("Sto eseguendo " + sProc)
        sResult = ""
        try:
            sOut = NormalizePath(self.out_xls_name)
            if self.is_xlsx(self.out_xls_name):
                if self.is_xlsx(self.in_sheet_file):
                    from openpyxl import load_workbook
                    wb_in = load_workbook(NormalizePath(self.in_sheet_file),
                                          data_only=True)
                    ws_in = None
                    for sName in wb_in.sheetnames:
                        if sName.lower() == self.in_sheet_name.lower():
                            ws_in = wb_in[sName]
                            break
                    if ws_in is None:
                        sResult = ("Non trovato " + self.in_sheet_name
                                   + " nel file " + self.in_sheet_file)
                        return ErrorProc(sResult, sProc)
                    self.num_righe_copiate, sErr = self.copy_openpyxl(
                        ws_in, self.ws_out)
                    if sErr:
                        return ErrorProc(sErr, sProc)
                else:
                    nRows, sErr = self.copy_xlrd_to_openpyxl(
                        self.in_sheet_file, self.in_sheet_name, self.ws_out)
                    if sErr:
                        return ErrorProc(sErr, sProc)
                    self.num_righe_copiate = nRows
                self.wb_out.save(sOut)
            else:
                booksheets, sErr = self.read_xls_book(self.out_xls_name)
                if sErr:
                    return ErrorProc(sErr, sProc)
                if self.is_xlsx(self.in_sheet_file):
                    from openpyxl import load_workbook
                    wb_in = load_workbook(NormalizePath(self.in_sheet_file),
                                          data_only=True)
                    ws_in = None
                    for sName in wb_in.sheetnames:
                        if sName.lower() == self.in_sheet_name.lower():
                            ws_in = wb_in[sName]
                            break
                    if ws_in is None:
                        sResult = ("Non trovato " + self.in_sheet_name
                                   + " nel file " + self.in_sheet_file)
                        return ErrorProc(sResult, sProc)
                    rows, merges, col_w, row_h, sErr = \
                        self.grid_from_openpyxl(ws_in)
                    if sErr:
                        return ErrorProc(sErr, sProc)
                else:
                    rows, merges, col_w, row_h, sErr = self.grid_from_xlrd(
                        self.in_sheet_file, self.in_sheet_name)
                    if sErr:
                        return ErrorProc(sErr, sProc)
                bReplaced = False
                booksheets_new = []
                for (sTitle, bRows, bMerges, bCw, bRh) in booksheets:
                    if sTitle.lower() == self.out_sheet_name.lower():
                        booksheets_new.append(
                            (sTitle, rows, merges, col_w, row_h))
                        bReplaced = True
                    else:
                        booksheets_new.append((sTitle, bRows, bMerges, bCw, bRh))
                if not bReplaced:
                    booksheets_new.append(
                        (self.out_sheet_name, rows, merges, col_w, row_h))
                sResult = self.write_xlwt(sOut, booksheets_new)
                if sResult:
                    return ErrorProc(sResult, sProc)
                self.num_righe_copiate = sum(
                    1 for r in rows if any(v is not None and v != "" for v in r))
            print("Copiate righe nello sheet: " + str(self.num_righe_copiate))
            print("File '" + self.out_xls_name + "' creato con successo.")
        except Exception as e:
            sResult = str(e)
        return ErrorProc(sResult, sProc)


def cmd_xls_tab_merge(dictJob: Dict[str, str]) -> str:
    """
    Esegue il comando XLS.TAB.MERGE in sequenza.

    arg_verify, engines_verify, read_list, header_read, header_check,
    read, exec. A fronte di errore interrompe e ritorna l'errore; a fronte di
    successo scrive come RETURN.VALUE il numero di righe scritte
    e dichiara l'output come RETURN.FILE.XLS.
    """
    sProc = "cmd_xls_tab_merge"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        oMerge = xls_merge()
        sResult = oMerge.arg_verify(dictJob)
        if not sResult:
            sResult = oMerge.engines_verify()
        if not sResult:
            sResult = oMerge.read_list()
        if not sResult:
            sResult = oMerge.header_read()
        if not sResult:
            sResult = oMerge.header_check()
        if not sResult:
            sResult = oMerge.read()
        if not sResult:
            sResult = oMerge.exec()
        if not sResult:
            sResult = jData.Return(
                "", str(oMerge.num_righe_scritte),
                {"XLS": oMerge.out_xls_name})
    except Exception as e:
        sResult = str(e)
    return ErrorProc(sResult, sProc)


def cmd_xls_append(dictJob: Dict[str, str]) -> str:
    """
    Esegue il comando XLS.SHEET.APPEND in sequenza.

    xls_append_args, xls_out_create, xls_sheet_create, xls_sheet_append.
    A fronte di errore interrompe e ritorna l'errore; a fronte di
    successo scrive come RETURN.VALUE il numero di righe copiate
    e dichiara l'output come RETURN.FILE.XLS.
    """
    sProc = "cmd_xls_append"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        oAppend = xls_append()
        sResult = oAppend.xls_append_args(dictJob)
        if not sResult:
            sResult = oAppend.xls_out_create()
        if not sResult:
            sResult = oAppend.xls_sheet_create()
        if not sResult:
            sResult = oAppend.xls_sheet_append()
        if not sResult:
            sResult = jData.Return(
                "", str(oAppend.num_righe_copiate),
                {"XLS": oAppend.out_xls_name})
    except Exception as e:
        sResult = str(e)
    return ErrorProc(sResult, sProc)


_DB_COMMANDS = ("DB.OPEN", "DB.CREATE", "DB.CLOSE",
                 "SQL.EXEC", "TAB.ZAP", "TAB.CREATE.JSON",
                 "DB.IMPORT", "DB.EXPORT")


def cbCommands(dictJob: Dict[str, str]) -> str:
    """
    Dispatcher dei comandi: instrada COMMAND alla funzione dedicata.

    Comandi DB.* in ntj_fact_db.py (vedi prompt_ntj_fact_db.md).
    """
    sProc = "cbCommands"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        sCommand = str(dictJob.get("COMMAND", "")).strip().upper()
        if sCommand == "XLS.TAB.MERGE":
            sResult = cmd_xls_tab_merge(dictJob)
        elif sCommand == "XLS.SHEET.APPEND":
            sResult = cmd_xls_append(dictJob)
        elif sCommand in _DB_COMMANDS:
            if _db_module_available:
                sResult = ntj_fact_db.cbCommands_db(dictJob)
            else:
                sResult = "Modulo ntj_fact_db non disponibile"
        elif not sCommand:
            sResult = "COMMAND non precisato"
        else:
            sResult = "COMMAND sconosciuto: " + sCommand
    except Exception as e:
        sResult = str(e)
    return ErrorProc(sResult, sProc)


def main() -> None:
    """
    Flusso ntJobsApp: Start, Run dei job, End con scrittura del .end,
    cleanup DB se ntj_fact_db.py e' stato usato.
    """
    sProc = "main"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        sResult = jData.Start()
        if sResult:
            jData.End(sResult)
            return
        sResult = jData.Run(cbCommands)
        jData.End(sResult)
        # Cleanup DB se il modulo e' stato usato
        if _db_module_available and ntj_fact_db.is_db_used():
            sClean = ntj_fact_db.db_cleanup_all()
            if sClean:
                print(sClean)
    except Exception as e:
        sResult = ErrorProc(str(e), sProc)
        print(sResult)
        sys.exit(1)


if __name__ == "__main__":
    main()
