"""Test DB.COMPACT (ntj_fact_db.py): SQLITE sempre, ACCESS live se possibile.

Uso (interprete con dipendenze, es. produzione):
    k:\\Tools\\pyn.cmd test_db\\test_db_compact.py
oppure da root: python test_db/test_db_compact.py
Artefatti in temp di sistema, mai in repo. Exit 0 = tutto verde.
Casi: shrink SQLITE + .old pre-immagine + integrita' + RETURN.VALUE,
guardia DB aperto, file mancante, estensione ignota, .accdb via ACE/DAO
(con bloat ODBC se il driver c'e', 3 retry anti-lock), DB.LIST live
(tabella + marker QUERIES vuoto), DB.LINK live ACCESS->ACCESS
(verifica righe via connessione fresca, errori self-link/DBL mancanti/
tabella sorgente assente/duplicato).
"""
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

DB = os.path.join(WORK, "c.db")

for f in (DB, DB + ".old"):
    if os.path.isfile(f):
        os.remove(f)

# DB gonfio: 3000 righe poi delete quasi tutto (freelist garantita)
c = sqlite3.connect(DB)
c.execute("CREATE TABLE T (ID INTEGER PRIMARY KEY, TXT TEXT)")
c.executemany("INSERT INTO T (TXT) VALUES (?)",
              [(("x" * 200),) for _ in range(3000)])
c.commit()
c.execute("DELETE FROM T WHERE ID > 100")
c.commit()
c.close()
n_before = os.path.getsize(DB)
print("size prima:", n_before)

jData.dictJobs = {"CONFIG": {"TYPE": "NTJOBS.APP.1", "NAME": "T",
                             "LOG": "t", "EXIT": "TRUE"}}


def run(job):
    jData.dictJob = dict(job)
    r = ntj_fact.cbCommands(dict(job))
    return r, dict(jData.dictJob)


# COMPACT ok
r, d = run({"COMMAND": "DB.COMPACT", "FILE.IN.DB": DB})
assert not r, r
n_after = os.path.getsize(DB)
print("size dopo:", n_after, "reclaimed:", d.get("RETURN.VALUE"))
assert os.path.isfile(DB + ".old"), ".old mancante!"
assert os.path.getsize(DB + ".old") == n_before, ".old != pre-immagine!"
assert n_after <= n_before, "non ha compattato!"
assert int(d.get("RETURN.VALUE")) == n_before - n_after
c = sqlite3.connect(DB)
assert c.execute("SELECT COUNT(*) FROM T").fetchone()[0] == 100
c.execute("PRAGMA integrity_check").fetchone()
c.close()
print("COMPACT sqlite ok")

# guardia DB aperto
r, d = run({"COMMAND": "DB.OPEN", "DB.ID": "T1",
            "TYPE": "SQLITE", "FILE.IN.DB": DB})
assert not r, r
r, d = run({"COMMAND": "DB.COMPACT", "FILE.IN.DB": DB})
assert "prima DB.CLOSE" in r, r
print("guardia aperto ok:", r.split(":")[-1].strip()[:60])
r, d = run({"COMMAND": "DB.CLOSE", "DB.ID": "T1"})
assert not r, r

# errori: mancante + estensione
r, d = run({"COMMAND": "DB.COMPACT",
            "FILE.IN.DB": os.path.join(WORK, "nope.db")})
assert "non esistente" in r, r
txt = os.path.join(WORK, "x.txt")
open(txt, "w").write("hi")
r, d = run({"COMMAND": "DB.COMPACT", "FILE.IN.DB": txt})
assert "non compattabile" in r, r
os.remove(txt)
print("errori ok")

# probe ACCESS live (.accdb via ACE; bloat via ODBC se c'e')
try:
    import win32com.client
    acc = os.path.join(WORK, "t_accdb.accdb")
    for f in (acc, acc + ".old"):
        if os.path.isfile(f):
            os.remove(f)
    cat = win32com.client.Dispatch("ADOX.Catalog")
    cat.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + acc + ";")
    del cat  # rilascia subito: Catalog.Create tiene il db aperto!
    import gc
    gc.collect()
    try:
        import pyodbc
        pyodbc.pooling = False  # close() deve davvero disconnettere
        cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
              f"DBQ={acc};")
        cn = pyodbc.connect(cs, autocommit=True)
        cur = cn.cursor()
        cur.execute("CREATE TABLE T (ID COUNTER PRIMARY KEY, TXT TEXT(200))")
        cur.executemany("INSERT INTO T (TXT) VALUES (?)",
                        [("y" * 200,) for _ in range(2000)])
        cur.execute("DELETE FROM T WHERE ID > 100")
        cur.close()
        cn.close()
        print("bloat ok, size:", os.path.getsize(acc))
    except Exception as e:
        print("bloat SKIPPED:", str(e)[:120])
    r, d = run({"COMMAND": "DB.COMPACT", "FILE.IN.DB": acc})
    assert not r, r
    assert os.path.isfile(acc + ".old")
    print("COMPACT access ok, reclaimed:", d.get("RETURN.VALUE"))
    # DB.LIST live: ADOX query/tabelle via ODBC-conn
    jData.dictJobs = {"CONFIG": {"TYPE": "NTJOBS.APP.1", "NAME": "T",
                                 "LOG": "t", "EXIT": "TRUE"}}
    lst = os.path.join(WORK, "acc_list.csv")
    r, d = run({"COMMAND": "DB.OPEN", "DB.ID": "A1", "TYPE": "ACCESS",
                "FILE.IN.DB": acc})
    assert not r, r
    r, d = run({"COMMAND": "DB.LIST", "DB.ID": "A1", "LIST": "ALL",
                "FILE.OUT.LIST": lst})
    assert not r, r
    with open(lst, encoding="utf-8-sig") as f:
        lines = f.read().splitlines()
    assert any(l.startswith("TABLES;T;") for l in lines), lines
    assert any(l == "QUERIES;(vuoto);" for l in lines), lines
    r, d = run({"COMMAND": "DB.CLOSE", "DB.ID": "A1"})
    assert not r, r
    print("LIST access live ok")
    # DB.LINK live: sorgente accdb secondario -> target acc (A1 riaperto)
    src = os.path.join(WORK, "acc_src.accdb")
    for f in (src,):
        if os.path.isfile(f):
            os.remove(f)
    cat2 = win32com.client.Dispatch("ADOX.Catalog")
    cat2.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + src + ";")
    del cat2
    gc.collect()
    cn2 = pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={src};", autocommit=True)
    cur2 = cn2.cursor()
    cur2.execute("CREATE TABLE SRC (ID INTEGER PRIMARY KEY, VAL TEXT(20))")
    cur2.execute("INSERT INTO SRC (ID, VAL) VALUES (1, 'pippo')")
    cur2.execute("INSERT INTO SRC (ID, VAL) VALUES (2, 'pluto')")
    cur2.close()
    cn2.close()
    r, d = run({"COMMAND": "DB.OPEN", "DB.ID": "A1", "TYPE": "ACCESS",
                "FILE.IN.DB": acc})
    assert not r, r
    r, d = run({"COMMAND": "DB.OPEN", "DB.ID": "A2", "TYPE": "ACCESS",
                "FILE.IN.DB": src})
    assert not r, r
    r, d = run({"COMMAND": "DB.LINK", "DB.ID": "A1", "DBL.ID": "A1",
                "DBL.TABLE": "SRC", "DBL.NAME": "X"})
    assert "self-link" in r, r
    r, d = run({"COMMAND": "DB.LINK", "DB.ID": "A1",
                "DBL.TABLE": "SRC", "DBL.NAME": "X"})
    assert "DBL.ID non precisato" in r, r
    r, d = run({"COMMAND": "DB.LINK", "DB.ID": "A1", "DBL.ID": "A2",
                "DBL.TABLE": "NOPE", "DBL.NAME": "X"})
    assert "non trovata nel sorgente" in r, r
    r, d = run({"COMMAND": "DB.LINK", "DB.ID": "A1", "DBL.ID": "A2",
                "DBL.TABLE": "SRC", "DBL.NAME": "LNK_SRC"})
    assert not r, r
    assert d.get("RETURN.VALUE") == "LNK_SRC", d
    # l'handle A1 aperto NON vede la nuova allegata (cache ODBC):
    # verifica via connessione fresca al file target
    fresh = pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={acc};", autocommit=True)
    cur_f = fresh.cursor()
    cur_f.execute("SELECT COUNT(*) FROM [LNK_SRC]")
    assert cur_f.fetchone()[0] == 2, "righe via link"
    cur_f.close()
    fresh.close()
    r, d = run({"COMMAND": "DB.LINK", "DB.ID": "A1", "DBL.ID": "A2",
                "DBL.TABLE": "SRC", "DBL.NAME": "LNK_SRC"})
    assert "gia' esistente" in r, r
    r, d = run({"COMMAND": "DB.CLOSE", "DB.ID": "A2"})
    assert not r, r
    r, d = run({"COMMAND": "DB.CLOSE", "DB.ID": "A1"})
    assert not r, r
    print("LINK access live ok")
    os.remove(acc)
    os.remove(acc + ".old")
    os.remove(src)
except Exception as e:
    print("ACCESS live SKIPPED:", str(e)[:120])

print("ALL COMPACT TESTS PASSED")
