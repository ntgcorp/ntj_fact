"""Test end-to-end comandi DB (ntj_fact_db.py) su SQLITE.

Uso (interprete con dipendenze, es. produzione):
    k:\\Tools\\pyn.cmd test_db\\test_db_cmds.py
oppure da root: python test_db/test_db_cmds.py
Artefatti in temp di sistema, mai in repo. Exit 0 = tutto verde.
Casi: OPEN/doppio-OPEN/ID-sconosciuto, TAB.CREATE.JSON/duplicata,
DB.IMPORT csv APPEND-REPLACE e json senza STRUCT (+errori),
SQL.EXEC SELECT/file/xor/DML + FILE.OUT.EXPORT csv/xls (+errori),
DB.EXPORT csv/json/xlsx per estensione (+errori), DB.INFO (con/senza
DB, errori), DB.LIST (TABLES/ALL/json/xlsx + errori), DB.LINK
(validazione target non ACCESS),
TAB.ZAP, CLOSE/doppia, dispatcher, cleanup finale.
"""
import json
import os
import sqlite3
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
WORK = os.path.join(tempfile.gettempdir(), "opencode")
os.makedirs(WORK, exist_ok=True)

import ntj_fact
import ntj_fact_db
from ntj_fact import jData

DB = os.path.join(WORK, "t.db")
STRUCT = os.path.join(WORK, "s.json")
CSVF = os.path.join(WORK, "d.csv")
JSONF = os.path.join(WORK, "d.json")
SQLF = os.path.join(WORK, "q.sql")

for f in (DB, DB + ".old",):
    if os.path.isfile(f):
        os.remove(f)

with open(STRUCT, "w", encoding="utf-8") as f:
    json.dump({"table": "UTENTI_TEST", "columns": [
        {"name": "ID", "type": "INTEGER", "pk": True, "auto": True,
         "null": False},
        {"name": "COD", "type": "TEXT", "size": 20, "null": False},
        {"name": "IMPORTO", "type": "DECIMAL", "size": "12,2"},
        {"name": "DATA", "type": "DATE"},
        {"name": "ATTIVO", "type": "BOOLEAN", "null": False}]}, f)

with open(CSVF, "w", encoding="utf-8") as f:
    f.write("COD;IMPORTO;DATA;ATTIVO\nA1;10.5;2026-01-02;1\n"
            "A2;;2026-03-04;0\nA3;7;;1\n")

with open(JSONF, "w", encoding="utf-8") as f:
    json.dump([{"COD": "B1", "IMPORTO": 1.5, "DATA": "2026-05-06",
                "ATTIVO": 1},
               {"COD": "B2", "ATTIVO": 0}], f)

with open(SQLF, "w", encoding="utf-8") as f:
    f.write("SELECT COD FROM UTENTI_TEST ORDER BY COD")

jData.dictJobs = {"CONFIG": {"TYPE": "NTJOBS.APP.1", "NAME": "T",
                             "LOG": "t", "EXIT": "TRUE"}}
OPEN_JOB = {"COMMAND": "DB.OPEN", "DB.ID": "T1",
            "TYPE": "SQLITE", "FILE.IN.DB": DB}


def run(job):
    jData.dictJob = dict(job)
    r = ntj_fact.cbCommands(dict(job))
    assert jData.dictJob is not None
    return r, dict(jData.dictJob)


def count():
    c = sqlite3.connect(DB)
    n = c.execute("SELECT COUNT(*) FROM UTENTI_TEST").fetchone()[0]
    c.close()
    return n


# OPEN + doppio OPEN + ID sconosciuto
r, d = run(dict(OPEN_JOB))
assert not r, r
assert d.get("RETURN.VALUE") == "T1", d
r, d = run(dict(OPEN_JOB))
assert "gia' aperto" in r, r
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "NOPE",
            "SQL": "SELECT 1"})
assert "non definito (manca DB.OPEN" in r, r
print("OPEN ok + doppio OPEN errore ok + ID sconosciuto ok")

# CREATE.JSON + duplicata
r, d = run({"COMMAND": "TAB.CREATE.JSON", "DB.ID": "T1",
            "FILE.IN.STRUCT": STRUCT})
assert not r, r
assert d.get("RETURN.VALUE") == "UTENTI_TEST", d
r, d = run({"COMMAND": "TAB.CREATE.JSON", "DB.ID": "T1",
            "FILE.IN.STRUCT": STRUCT})
assert "gia' esistente" in r, r
print("CREATE.JSON ok + duplicata errore ok")

# IMPORT csv APPEND x2 poi REPLACE
r, d = run({"COMMAND": "DB.IMPORT", "DB.ID": "T1",
            "FILE.IN.STRUCT": STRUCT, "FILE.IN.CSV": CSVF})
assert not r, r
assert d.get("RETURN.VALUE") == "3", d
assert count() == 3
r, d = run({"COMMAND": "DB.IMPORT", "DB.ID": "T1",
            "FILE.IN.STRUCT": STRUCT, "FILE.IN.CSV": CSVF})
assert count() == 6
r, d = run({"COMMAND": "DB.IMPORT", "DB.ID": "T1",
            "FILE.IN.STRUCT": STRUCT, "FILE.IN.CSV": CSVF,
            "MODE": "REPLACE"})
assert not r, r
assert count() == 3
print("IMPORT csv APPEND/REPLACE ok")

# IMPORT json senza STRUCT (introspezione) + TABLE mancante
r, d = run({"COMMAND": "DB.IMPORT", "DB.ID": "T1",
            "FILE.IN.JSON": JSONF, "TABLE": "UTENTI_TEST"})
assert not r, r
assert count() == 5, count()
r, d = run({"COMMAND": "DB.IMPORT", "DB.ID": "T1",
            "FILE.IN.JSON": JSONF})
assert "TABLE obbligatorio" in r, r
r, d = run({"COMMAND": "DB.IMPORT", "DB.ID": "T1",
            "FILE.IN.STRUCT": STRUCT, "FILE.IN.CSV": CSVF, "MODE": "XX"})
assert "MODE non valido" in r, r
print("IMPORT json no-STRUCT ok + errori ok")

# SQL.EXEC SELECT + FILE + entrambi + DML
XLSX = os.path.join(WORK, "out.xlsx")
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1",
            "SQL": "SELECT COD FROM UTENTI_TEST ORDER BY COD",
            "FILE.OUT.XLS": XLSX})
assert not r, r
assert d.get("RETURN.VALUE") == "5", d
assert d.get("RETURN.FILE.XLS") == "out.xlsx", d
assert os.path.isfile(XLSX)
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1", "FILE.IN.SQL": SQLF})
assert not r, r
assert d.get("RETURN.VALUE") == "5", d
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1", "SQL": "SELECT 1",
            "FILE.IN.SQL": SQLF})
assert "xor" in r, r
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1",
            "SQL": "DELETE FROM UTENTI_TEST WHERE COD='ZZZ'"})
assert not r, r
assert d.get("RETURN.VALUE") == "0", d
print("SQL.EXEC ok + errori ok")

# SQL.EXEC FILE.OUT.EXPORT: csv + xls + errori
ECVS = os.path.join(WORK, "sqlexec.csv")
EXLS = os.path.join(WORK, "sqlexec.xls")
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1",
            "SQL": "SELECT COD, ATTIVO FROM UTENTI_TEST ORDER BY COD",
            "FILE.OUT.EXPORT": ECVS})
assert not r, r
assert d.get("RETURN.VALUE") == "5" and \
    d.get("RETURN.FILE.EXPORT") == "sqlexec.csv", d
with open(ECVS, encoding="utf-8-sig") as f:
    lines = f.read().splitlines()
assert lines[0] == "COD;ATTIVO" and len(lines) == 6, lines
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1",
            "SQL": "SELECT COD FROM UTENTI_TEST ORDER BY COD",
            "FILE.OUT.EXPORT": EXLS})
assert not r, r
assert d.get("RETURN.FILE.EXPORT") == "sqlexec.xls", d
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1", "SQL": "DELETE FROM UTENTI_TEST",
            "FILE.OUT.EXPORT": ECVS})
assert "solo per SELECT" in r, r
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1", "SQL": "SELECT 1",
            "FILE.OUT.XLS": XLSX, "FILE.OUT.EXPORT": ECVS})
assert "xor" in r, r
TXT = os.path.join(WORK, "out.txt")
open(TXT, "w").write("x")
r, d = run({"COMMAND": "SQL.EXEC", "DB.ID": "T1", "SQL": "SELECT 1",
            "FILE.OUT.EXPORT": TXT})
assert "non valida" in r, r
os.remove(TXT)
print("SQL.EXEC EXPORT csv/xls ok + errori ok")

# DB.EXPORT csv/json/xlsx per estensione + errori
ECSV = os.path.join(WORK, "exp.csv")
EJSON = os.path.join(WORK, "exp.json")
EXLSX = os.path.join(WORK, "exp.xlsx")
r, d = run({"COMMAND": "DB.EXPORT", "DB.ID": "T1",
            "OUT.QUERY": "SELECT COD, ATTIVO FROM UTENTI_TEST ORDER BY COD",
            "FILE.OUT.QUERY": ECSV})
assert not r, r
assert d.get("RETURN.VALUE") == "5", d
assert d.get("RETURN.FILE.QUERY") == "exp.csv", d
with open(ECSV, encoding="utf-8-sig") as f:
    lines = f.read().splitlines()
assert lines[0] == "COD;ATTIVO", lines[0]
assert len(lines) == 6, len(lines)
r, d = run({"COMMAND": "DB.EXPORT", "DB.ID": "T1",
            "OUT.QUERY": "SELECT COD FROM UTENTI_TEST ORDER BY COD",
            "FILE.OUT.QUERY": EJSON})
assert not r, r
with open(EJSON, encoding="utf-8") as f:
    jd = json.load(f)
assert isinstance(jd, list) and len(jd) == 5 and jd[0]["COD"] == "A1", jd
r, d = run({"COMMAND": "DB.EXPORT", "DB.ID": "T1",
            "OUT.QUERY": "SELECT COD FROM UTENTI_TEST ORDER BY COD",
            "FILE.OUT.QUERY": EXLSX})
assert not r, r
assert os.path.isfile(EXLSX)
r, d = run({"COMMAND": "DB.EXPORT", "DB.ID": "T1",
            "OUT.QUERY": "DELETE FROM UTENTI_TEST",
            "FILE.OUT.QUERY": ECSV})
assert "deve essere una SELECT" in r, r
ETXT = os.path.join(WORK, "exp.txt")
r, d = run({"COMMAND": "DB.EXPORT", "DB.ID": "T1",
            "OUT.QUERY": "SELECT 1", "FILE.OUT.QUERY": ETXT})
assert "non supportata" in r, r
print("DB.EXPORT csv/json/xlsx ok + errori ok")

# DB.INFO: con DB aperto + senza DB.ID + errori (prima dello ZAP)
INFO = os.path.join(WORK, "info.txt")
r, d = run({"COMMAND": "DB.INFO", "DB.ID": "T1", "FILE.OUT.INFO": INFO})
assert not r, r
assert d.get("RETURN.FILE.INFO") == "info.txt", d
with open(INFO, encoding="utf-8") as f:
    txt = f.read()
assert "DB <T1> (SQLITE)" in txt, txt[:400]
assert "UTENTI_TEST: 5 righe" in txt, txt[:400]
assert "SQLite" in txt, txt[:400]
assert "DRIVER ODBC" in txt, txt[:400]
INFO2 = os.path.join(WORK, "info_nodb.txt")
r, d = run({"COMMAND": "DB.INFO", "FILE.OUT.INFO": INFO2})
assert not r, r
with open(INFO2, encoding="utf-8") as f:
    txt2 = f.read()
assert "DB <" not in txt2 and "DRIVER ODBC" in txt2
INFO3 = os.path.join(WORK, "info.doc")
r, d = run({"COMMAND": "DB.INFO", "FILE.OUT.INFO": INFO3})
assert "deve finire con .txt" in r, r
r, d = run({"COMMAND": "DB.INFO", "FILE.OUT.INFO": "   "})
assert "non precisato" in r, r
INFOSEQ = os.path.join(WORK, "info_seq.txt")
r, d = run({"COMMAND": "DB.INFO", "DB.ID": "NOPE",
            "FILE.OUT.INFO": INFOSEQ})
assert "non definito" in r, r

# DB.LIST: TABLES + ALL + formati + errori (prima dello ZAP)
LCVS = os.path.join(WORK, "lst.csv")
r, d = run({"COMMAND": "DB.LIST", "DB.ID": "T1", "LIST": "TABLES",
            "FILE.OUT.LIST": LCVS})
assert not r, r
assert d.get("RETURN.FILE.LIST") == "lst.csv", d
with open(LCVS, encoding="utf-8-sig") as f:
    lines = f.read().splitlines()
assert lines[0] == "KIND;NAME;INFO", lines[0]
assert any(l == "TABLES;UTENTI_TEST;" for l in lines), lines
LALL = os.path.join(WORK, "lstall.csv")
r, d = run({"COMMAND": "DB.LIST", "DB.ID": "T1",
            "FILE.OUT.LIST": LALL})
assert not r, r
with open(LALL, encoding="utf-8-sig") as f:
    lines = f.read().splitlines()
kinds = {l.split(";")[0] for l in lines[1:]}
assert "TABLES" in kinds and "PKFK" in kinds, kinds
assert "PK UTENTI_TEST" in ";".join(lines), lines
assert not any(s.startswith("QUERIES;") for s in lines), lines
LJSON = os.path.join(WORK, "lst.json")
r, d = run({"COMMAND": "DB.LIST", "DB.ID": "T1", "LIST": "PKFK",
            "FILE.OUT.LIST": LJSON})
assert not r, r
with open(LJSON, encoding="utf-8") as f:
    jl = json.load(f)
assert isinstance(jl, list) and jl[0]["KIND"] == "PKFK" and \
    jl[0]["NAME"].startswith("PK "), jl
LXLSX = os.path.join(WORK, "lst.xlsx")
r, d = run({"COMMAND": "DB.LIST", "DB.ID": "T1", "LIST": "INDEXES",
            "FILE.OUT.LIST": LXLSX})
assert not r, r
assert os.path.isfile(LXLSX)
r, d = run({"COMMAND": "DB.LIST", "DB.ID": "T1", "LIST": "XX",
            "FILE.OUT.LIST": LCVS})
assert "LIST non valida" in r, r
r, d = run({"COMMAND": "DB.LIST", "DB.ID": "T1",
            "FILE.OUT.LIST": "lst.txt"})
assert "non supportata" in r, r
r, d = run({"COMMAND": "DB.LIST", "DB.ID": "NOPE",
            "FILE.OUT.LIST": LCVS})
assert "non definito" in r, r
print("DB.LIST ok + errori ok")

# DB.LINK: validazione (con target SQLITE esce subito "deve essere ACCESS")
r, d = run({"COMMAND": "DB.LINK", "DB.ID": "T1", "DBL.ID": "T1",
            "DBL.TABLE": "UTENTI_TEST", "DBL.NAME": "LNK"})
assert "deve essere di tipo ACCESS" in r, r
r, d = run({"COMMAND": "DB.LINK", "DB.ID": "T1"})
assert "deve essere di tipo ACCESS" in r, r
print("DB.LINK validazione ok")

# ZAP + CLOSE + CLOSE-doppia + cleanup
r, d = run({"COMMAND": "TAB.ZAP", "DB.ID": "T1", "TABLE": "UTENTI_TEST"})
assert not r, r
assert d.get("RETURN.VALUE") == "5", d
assert count() == 0
r, d = run({"COMMAND": "DB.CLOSE", "DB.ID": "T1"})
assert not r, r
r, d = run({"COMMAND": "DB.CLOSE", "DB.ID": "T1"})
assert "non aperto" in r, r
assert ntj_fact_db.is_db_used() is True
print("ZAP/CLOSE ok + errori ok")

# dispatcher: sconosciuto e vuoto
r = ntj_fact.cbCommands({"COMMAND": "FOO"})
assert "sconosciuto" in r, r
r = ntj_fact.cbCommands({"COMMAND": ""})
assert "non precisato" in r, r
print("dispatcher ok")

# cleanup finale con DB aperto
r, d = run(dict(OPEN_JOB))
assert not r, r
assert ntj_fact_db.db_cleanup_all() == ""
assert ntj_fact_db.dictDb["T1"]["STATUS"] == "CLOSE"
print("cleanup ok")

# SQLSERVER: validazione parametri (connessione live non testabile qui)
r, d = run({"COMMAND": "DB.OPEN", "DB.ID": "S1", "TYPE": "SQLSERVER"})
assert "DBNAME mancante" in r, r
err, sid, scfg = ntj_fact_db._read_db_params(
    {"DB.ID": "S1", "TYPE": "SQLSERVER", "DBNAME": "P"})
assert not err, err
print("SQLSERVER param validation ok")
print("ALL DB TESTS PASSED")
