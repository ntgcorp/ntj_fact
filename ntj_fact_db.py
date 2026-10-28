"""
ntj_fact_db.py - Gruppo comandi database per ntj_fact (ntJobsApp).

Specifica: prompt_ntj_fact_db.md. Comandi: DB.OPEN, DB.CREATE, DB.CLOSE,
SQL.EXEC, TAB.ZAP, TAB.CREATE.JSON, DB.IMPORT, DB.EXPORT, DB.COMPACT.
Tutti i comandi funzionano
per TYPE: SQLITE, ACCESS, SQLSERVER, ORACLE, POSTGRES, MARIADB
(driver specifici).

Regola nomi file: input aperti = FILE.IN.*, output creati = FILE.OUT.*.

Uso: importato da ntj_fact.py; espone cbCommands_db(dictJob) e
db_cleanup_all() da chiamare a fine run se usato.
"""

import os
import sys
import json
import csv
import re
import time
import shutil
from typing import Dict, List, Any, Optional, Tuple

# Flag: True se almeno un comando DB e' stato eseguito in questo run
_db_used = False

# Dizionario globale dei database: {ID: {TYPE, ..., HANDLE, STATUS}}.
# Popolato da DB.OPEN (e DB.CREATE, chiuso) dai parametri del job.
# STATUS vale "OPEN" o "CLOSE". Nessuna sezione [DB_*] in INI.
dictDb: Dict[str, Dict[str, Any]] = {}


def _jdata() -> Any:
    """Ritorna jData RUNNING (quello di __main__, non una copia fresca).

    Attenzione: se ntj_fact.py gira come script, l'import qui dentro
    ricaricherebbe il modulo con un jData vuoto separato. Si usa quindi
    jData di __main__ quando disponibile, con fallback all'import.
    """
    try:
        import sys as _sys
        m = _sys.modules.get("__main__")
        jd = getattr(m, "jData", None)
        if jd is not None and hasattr(jd, "dictJobs"):
            return jd
    except Exception:
        pass
    try:
        from ntj_fact import jData
        return jData
    except Exception:
        return None


def _log_info(sMsg: str) -> None:
    """Video + file log via jData se disponibile, altrimenti print."""
    try:
        jd = _jdata()
        if jd is not None:
            jd.Log1(sMsg)
            return
    except Exception:
        pass
    print(sMsg)


def _do_return(sValue: str, dictFiles: Optional[Dict[str, str]] = None) -> str:
    """Scrive RETURN.VALUE (+RETURN.FILE.*) sul dictJob corrente."""
    try:
        jd = _jdata()
        if jd is None:
            return ""
        if dictFiles:
            return jd.Return("", sValue, dictFiles)
        return jd.Return("", sValue)
    except Exception as e:
        return str(e)


def _error_proc(sResult: str, sProc: str) -> str:
    """Wrapper ErrorProc (ntj_fact lo riwrappera' col suo)."""
    if sResult:
        return f"{sProc}: Errore {sResult}"
    return sResult


# =============================================================================
# DRIVER ABSTRACTION (placeholder per tutti i TYPE)
# =============================================================================

def _check_driver(db_type: str) -> str:
    """Verifica dipendenze per TYPE. Ritorna "" o errore."""
    sProc = "_check_driver"
    try:
        if db_type == "SQLITE":
            import sqlite3  # noqa: F401
        elif db_type in ("ACCESS", "SQLSERVER"):
            import pyodbc  # noqa: F401
        elif db_type == "ORACLE":
            import oracledb  # noqa: F401
        elif db_type == "POSTGRES":
            import psycopg  # noqa: F401
        elif db_type == "MARIADB":
            import pymysql  # noqa: F401
        else:
            return f"TYPE sconosciuto: {db_type}"
        return ""
    except ImportError as e:
        return f"dipendenza per {db_type} non installata: {e}"
    except Exception as e:
        return str(e)


def _db_connect(db_id: str, cfg: Dict[str, str]) -> Any:
    """Apre connessione per TYPE. Ritorna conn o solleva eccezione."""
    sProc = "_db_connect"
    db_type = cfg.get("TYPE", "")
    timeout = int(cfg.get("TIMEOUT", "30"))
    if db_type == "SQLITE":
        import sqlite3
        path = cfg.get("FILE.IN.DB", "")
        if not path:
            raise ValueError("FILE.IN.DB mancante per SQLITE")
        # READONLY gestito da chi chiama
        conn = sqlite3.connect(path, timeout=timeout)
        return conn
    elif db_type == "ACCESS":
        import pyodbc
        # Path assoluto: il driver ODBC non risolve i relativi in CWD.
        # Pooling off: close() deve davvero disconnettere, altrimenti
        # il file resta lockato (es. per DB.COMPACT successivo).
        pyodbc.pooling = False
        path = os.path.abspath(cfg.get("FILE.IN.DB", ""))
        driver = cfg.get("DRIVER", "Microsoft Access Driver (*.mdb, *.accdb)")
        uid = cfg.get("UID", "")
        pwd = cfg.get("PASSWORD", "")
        cs = f"DRIVER={{{driver}}};DBQ={path};"
        if uid:
            cs += f"UID={uid};"
        if pwd:
            cs += f"PWD={pwd};"
        conn = pyodbc.connect(cs, timeout=timeout)
        return conn
    elif db_type == "SQLSERVER":
        import pyodbc
        host = cfg.get("HOST", "localhost")
        port = cfg.get("PORT", "1433")
        dbname = cfg.get("DBNAME", "master")
        user = cfg.get("USER", "")
        pwd = cfg.get("PASSWORD", "")
        driver = cfg.get("DRIVER", "ODBC Driver 17 for SQL Server")
        cs = (f"DRIVER={{{driver}}};SERVER={host},{port};"
              f"DATABASE={dbname};")
        if user:
            cs += f"UID={user};PWD={pwd};"
        else:
            # Senza USER: autenticazione Windows dell'utente corrente
            cs += "Trusted_Connection=yes;"
        cs += "TrustServerCertificate=yes;"
        conn = pyodbc.connect(cs, timeout=timeout)
        return conn
    elif db_type == "ORACLE":
        import oracledb
        host = cfg.get("HOST", "")
        port = cfg.get("PORT", "1521")
        service = cfg.get("SERVICE", "")
        tns = cfg.get("TNSNAME", "")
        user = cfg.get("USER", "")
        pwd = cfg.get("PASSWORD", "")
        if tns:
            dsn = tns
        elif host and service:
            dsn = oracledb.makedsn(host, int(port), service_name=service)
        else:
            raise ValueError("HOST+SERVICE o TNSNAME obbligatori per ORACLE")
        conn = oracledb.connect(user=user, password=pwd, dsn=dsn,
                                timeout=timeout)
        return conn
    elif db_type == "POSTGRES":
        import psycopg
        host = cfg.get("HOST", "localhost")
        port = cfg.get("PORT", "5432")
        dbname = cfg.get("DBNAME", "")
        user = cfg.get("USER", "")
        pwd = cfg.get("PASSWORD", "")
        schema = cfg.get("SCHEMA", "public")
        sslmode = cfg.get("SSLMODE", "prefer")
        conn = psycopg.connect(
            host=host, port=port, dbname=dbname, user=user,
            password=pwd, sslmode=sslmode,
            options=f"-c search_path={schema}" if schema else None,
            connect_timeout=timeout)
        return conn
    elif db_type == "MARIADB":
        import pymysql
        host = cfg.get("HOST", "localhost")
        port = int(cfg.get("PORT", "3306"))
        dbname = cfg.get("DBNAME", "")
        user = cfg.get("USER", "")
        pwd = cfg.get("PASSWORD", "")
        charset = cfg.get("CHARSET", "utf8mb4")
        conn = pymysql.connect(
            host=host, port=port, database=dbname, user=user,
            password=pwd, charset=charset,
            connect_timeout=timeout)
        return conn
    else:
        raise ValueError(f"TYPE non supportato: {db_type}")


def _quote_ident(db_type: str, name: str) -> str:
    """Quota identificatore per TYPE."""
    if not re.match(r"^[A-Za-z0-9_]+$", name):
        raise ValueError(f"Nome non valido: {name}")
    if db_type in ("SQLITE", "ORACLE", "POSTGRES"):
        return f'"{name}"'
    elif db_type in ("ACCESS", "SQLSERVER"):
        return f"[{name}]"
    elif db_type == "MARIADB":
        return f"`{name}`"
    else:
        return name


def _placeholder(db_type: str, idx: int) -> str:
    """Restituisce placeholder per TYPE (idx 1-based per ORACLE)."""
    if db_type in ("SQLITE", "ACCESS", "SQLSERVER"):
        return "?"
    elif db_type == "ORACLE":
        return f":{idx}"
    elif db_type in ("POSTGRES", "MARIADB"):
        return "%s"
    else:
        return "?"


def _tab_exists(conn: Any, db_type: str, table: str) -> bool:
    """Verifica esistenza tabella per TYPE."""
    cur = conn.cursor()
    try:
        if db_type == "SQLITE":
            cur.execute(
                "SELECT COUNT(*) FROM sqlite_master "
                "WHERE type='table' AND name=?", (table,))
        elif db_type == "ACCESS":
            # Probe con SELECT COUNT(*) (MSysObjects spesso negato).
            # Tabella mancante = eccezione driver -> False (non esiste).
            try:
                q = f"SELECT COUNT(*) FROM {_quote_ident(db_type, table)}"
                cur.execute(q)
            except Exception:
                try:
                    conn.rollback()
                except Exception:
                    pass
                return False
        elif db_type == "ORACLE":
            cur.execute(
                "SELECT COUNT(*) FROM user_tables "
                "WHERE table_name = UPPER(?)", (table,))
        elif db_type == "POSTGRES":
            cur.execute(
                "SELECT COUNT(*) FROM information_schema.tables "
                "WHERE table_name = %s", (table,))
        elif db_type == "SQLSERVER":
            cur.execute(
                "SELECT COUNT(*) FROM information_schema.tables "
                "WHERE table_name = ?", (table,))
        elif db_type == "MARIADB":
            cur.execute(
                "SELECT COUNT(*) FROM information_schema.tables "
                "WHERE table_name = %s", (table,))
        else:
            return False
        row = cur.fetchone()
        return bool(row and row[0])
    finally:
        cur.close()


def _type_to_ddl(db_type: str, col: Dict[str, Any]) -> str:
    """Traduce tipo neutrale → DDL per TYPE."""
    t_neu = str(col.get("type", "TEXT")).upper()
    size = col.get("size")
    pk = col.get("pk", False)
    auto = col.get("auto", False)
    nullable = col.get("null", True)
    if db_type == "SQLITE":
        if t_neu == "INTEGER" and pk and auto:
            base = "INTEGER PRIMARY KEY AUTOINCREMENT"
        elif t_neu == "INTEGER":
            base = "INTEGER"
        elif t_neu == "TEXT":
            base = "TEXT"
        elif t_neu.startswith("DECIMAL"):
            base = "NUMERIC"
        elif t_neu in ("DATE", "DATETIME"):
            base = "TEXT"
        elif t_neu == "BOOLEAN":
            base = "INTEGER"
        else:
            base = "TEXT"
    elif db_type == "ACCESS":
        if t_neu == "INTEGER" and pk and auto:
            base = "COUNTER PRIMARY KEY"
        elif t_neu == "INTEGER":
            base = "LONG"
        elif t_neu == "TEXT":
            base = f"TEXT({size})" if size else "TEXT(255)"
        elif t_neu.startswith("DECIMAL"):
            base = f"DECIMAL({size})" if size else "DECIMAL(18,2)"
        elif t_neu in ("DATE", "DATETIME"):
            base = "DATETIME"
        elif t_neu == "BOOLEAN":
            base = "YESNO"
        else:
            base = "TEXT(255)"
    elif db_type == "ORACLE":
        if t_neu == "INTEGER" and pk and auto:
            base = "NUMBER GENERATED ALWAYS AS IDENTITY PRIMARY KEY"
        elif t_neu == "INTEGER":
            base = "NUMBER(10)"
        elif t_neu == "TEXT":
            base = f"VARCHAR2({size})" if size else "VARCHAR2(255)"
        elif t_neu.startswith("DECIMAL"):
            base = f"NUMBER({size})" if size else "NUMBER(18,2)"
        elif t_neu in ("DATE", "DATETIME"):
            base = "DATE" if t_neu == "DATE" else "TIMESTAMP"
        elif t_neu == "BOOLEAN":
            base = "NUMBER(1)"
        else:
            base = "VARCHAR2(255)"
    elif db_type == "POSTGRES":
        if t_neu == "INTEGER" and pk and auto:
            base = "INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY"
        elif t_neu == "INTEGER":
            base = "INTEGER"
        elif t_neu == "TEXT":
            base = f"VARCHAR({size})" if size else "TEXT"
        elif t_neu.startswith("DECIMAL"):
            base = f"NUMERIC({size})" if size else "NUMERIC(18,2)"
        elif t_neu == "DATE":
            base = "DATE"
        elif t_neu == "DATETIME":
            base = "TIMESTAMP"
        elif t_neu == "BOOLEAN":
            base = "BOOLEAN"
        else:
            base = "TEXT"
    elif db_type == "SQLSERVER":
        if t_neu == "INTEGER" and pk and auto:
            base = "INT IDENTITY(1,1) PRIMARY KEY"
        elif t_neu == "INTEGER":
            base = "INT"
        elif t_neu == "TEXT":
            base = f"VARCHAR({size})" if size else "VARCHAR(255)"
        elif t_neu.startswith("DECIMAL"):
            base = f"DECIMAL({size})" if size else "DECIMAL(18,2)"
        elif t_neu == "DATE":
            base = "DATE"
        elif t_neu == "DATETIME":
            base = "DATETIME2"
        elif t_neu == "BOOLEAN":
            base = "BIT"
        else:
            base = "VARCHAR(255)"
    elif db_type == "MARIADB":
        if t_neu == "INTEGER" and pk and auto:
            base = "INT AUTO_INCREMENT PRIMARY KEY"
        elif t_neu == "INTEGER":
            base = "INT"
        elif t_neu == "TEXT":
            base = f"VARCHAR({size})" if size else "VARCHAR(255)"
        elif t_neu.startswith("DECIMAL"):
            base = f"DECIMAL({size})" if size else "DECIMAL(18,2)"
        elif t_neu == "DATE":
            base = "DATE"
        elif t_neu == "DATETIME":
            base = "DATETIME"
        elif t_neu == "BOOLEAN":
            base = "TINYINT(1)"
        else:
            base = "VARCHAR(255)"
    else:
        base = "TEXT"
    if not (pk and auto):  # gia' incluso in alcuni casi
        if not nullable:
            base += " NOT NULL"
    return base


# =============================================================================
# COMANDI DB.*
# =============================================================================

_DB_TYPES = ("SQLITE", "ACCESS", "ORACLE", "POSTGRES", "MARIADB",
             "SQLSERVER")


def _db_cfg(db_id: str, need_open: bool) -> Tuple[str, Optional[Dict[str, Any]]]:
    """Risolve ID in dictDb (popolato da DB.OPEN). Ritorna (errore, cfg|None)."""
    db_id = str(db_id or "").strip().upper()
    if not db_id:
        return ("DB.ID non precisato", None)
    if db_id not in dictDb:
        return (f"DB <{db_id}> non definito (manca DB.OPEN con parametri)", None)
    cfg = dictDb[db_id]
    if need_open and cfg.get("STATUS") != "OPEN":
        return (f"DB <{db_id}> non aperto (manca DB.OPEN)", None)
    return ("", cfg)


def _read_db_params(dictJob: Dict[str, str]
                    ) -> Tuple[str, str, Dict[str, Any]]:
    """Legge+valida DB.ID/TYPE/parametri da job. Ritorna (errore, id, cfg)."""
    sProc = "_read_db_params"
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        if not db_id:
            return ("DB.ID non precisato", "", {})
        if not re.match(r"^[A-Z0-9_]+$", db_id):
            return (f"DB.ID non valido: {db_id}", "", {})
        db_type = str(dictJob.get("TYPE", "")).strip().upper()
        if not db_type:
            return ("TYPE non precisato", "", {})
        if db_type not in _DB_TYPES:
            return (f"TYPE sconosciuto: {db_type}", "", {})
        cfg = {"TYPE": db_type}  # type: Dict[str, Any]
        for k in ("TIMEOUT", "READONLY", "DRIVER", "UID", "PASSWORD",
                  "HOST", "PORT", "SERVICE", "SIDTYPE", "TNSNAME",
                  "DBNAME", "USER", "SCHEMA", "SSLMODE", "CHARSET",
                  "FILE.IN.DB"):
            v = dictJob.get(k, "")
            if isinstance(v, str):
                v = v.strip()
            if v != "":
                cfg[k] = v
        if "TIMEOUT" not in cfg:
            cfg["TIMEOUT"] = "30"
        if db_type in ("SQLITE", "ACCESS"):
            if not str(cfg.get("FILE.IN.DB", "")).strip():
                return (f"FILE.IN.DB mancante per {db_type} (deve esistere)",
                        "", {})
        if db_type == "ORACLE":
            if not (str(cfg.get("TNSNAME", "")).strip() or
                    (str(cfg.get("HOST", "")).strip() and
                     str(cfg.get("SERVICE", "")).strip())):
                return ("HOST+SERVICE o TNSNAME obbligatori per ORACLE",
                        "", {})
            if not str(cfg.get("USER", "")).strip():
                return ("USER mancante per ORACLE", "", {})
        if db_type in ("POSTGRES", "MARIADB"):
            if not str(cfg.get("DBNAME", "")).strip():
                return (f"DBNAME mancante per {db_type}", "", {})
            if not str(cfg.get("USER", "")).strip():
                return (f"USER mancante per {db_type}", "", {})
        if db_type == "SQLSERVER":
            if not str(cfg.get("DBNAME", "")).strip():
                return ("DBNAME mancante per SQLSERVER", "", {})
        return ("", db_id, cfg)
    except Exception as e:
        return (f"{sProc}: Errore {e}", "", {})

def cmd_db_open(dictJob: Dict[str, str]) -> str:
    """DB.OPEN: DB.ID + TYPE + parametri per TYPE (vedi prompt_ntj_fact_db.md).

    Registra il DB in dictDb con flag STATUS="OPEN" (piu' DB aperti insieme).
    """
    global _db_used, dictDb
    sProc = "cmd_db_open"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        sResult, db_id, cfg = _read_db_params(dictJob)
        if sResult:
            return _error_proc(sResult, sProc)
        if db_id in dictDb and dictDb[db_id].get("STATUS") == "OPEN":
            return _error_proc(f"DB <{db_id}> gia' aperto", sProc)
        db_type = cfg.get("TYPE", "")
        sResult = _check_driver(db_type)
        if sResult:
            return _error_proc(sResult, sProc)
        readonly = str(cfg.get("READONLY", "FALSE")).upper() == "TRUE"
        if db_type == "SQLITE" and readonly:
            import sqlite3
            conn = sqlite3.connect(
                f"file:{cfg['FILE.IN.DB']}?mode=ro", uri=True,
                timeout=int(cfg.get("TIMEOUT", "30")))
        else:
            conn = _db_connect(db_id, cfg)
        cfg["HANDLE"] = conn
        cfg["STATUS"] = "OPEN"
        dictDb[db_id] = cfg
        _db_used = True
        _log_info(f"DB <{db_id}> aperto ({db_type})")
        sResult = _do_return(db_id)
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def cmd_db_create(dictJob: Dict[str, str]) -> str:
    """DB.CREATE: DB.ID + TYPE + parametri per TYPE (come DB.OPEN).

    Crea il contenitore se manca e registra il DB chiuso in dictDb.
    """
    global _db_used
    sProc = "cmd_db_create"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        sResult, db_id, cfg = _read_db_params(dictJob)
        if sResult:
            return _error_proc(sResult, sProc)
        db_type = cfg.get("TYPE", "")
        sResult = _check_driver(db_type)
        if sResult:
            return _error_proc(sResult, sProc)
        # Implementazione per TYPE (FILE.IN.DB verificato da Start in CWD)
        if db_type == "SQLITE":
            import sqlite3
            path = cfg.get("FILE.IN.DB", "")
            parent = os.path.dirname(path)
            if parent and not os.path.isdir(parent):
                os.makedirs(parent, exist_ok=True)
            # Creazione file: apro e chiudo (idempotente se esiste)
            bExisted = os.path.isfile(path)
            conn = sqlite3.connect(path, timeout=int(cfg.get("TIMEOUT", "30")))
            conn.close()
            if bExisted:
                _log_info(f"DB <{db_id}> esiste gia' (file {path})")
            else:
                _log_info(f"DB <{db_id}> creato (file {path})")
        elif db_type == "ACCESS":
            # Richiede ADOX/COM (solo Windows + Access Engine).
            # Esiste gia' = INFO idempotente (ADOX.Create fallirebbe).
            acc_path = cfg.get("FILE.IN.DB", "")
            if acc_path and os.path.isfile(acc_path):
                _log_info(f"DB <{db_id}> esiste gia' (ACCESS)")
            else:
                try:
                    import win32com.client  # noqa: F401
                except ImportError:
                    return _error_proc(
                        "ACCESS CREATE richiede pywin32 (pip install pywin32)",
                        sProc)
                cat = win32com.client.Dispatch("ADOX.Catalog")
                conn_str = (
                    "Provider=Microsoft.ACE.OLEDB.12.0;"
                    f"Data Source={acc_path};")
                cat.Create(conn_str)
                _log_info(f"DB <{db_id}> creato (ACCESS)")
        elif db_type in ("ORACLE", "POSTGRES", "MARIADB", "SQLSERVER"):
            # Richiede utenza con privilegio CREATEDB / DBA
            # Connessione amministrativa temporanea
            tmp_cfg = cfg.copy()
            tmp_cfg["DBNAME"] = ""  # connessione al server
            conn = _db_connect(db_id, tmp_cfg)
            try:
                if getattr(conn, "autocommit", None) is False:
                    conn.autocommit = True  # CREATE DATABASE fuori txn
            except Exception:
                pass
            dbname = cfg.get("DBNAME", f"DB_{db_id}")
            cur = conn.cursor()
            try:
                if db_type == "POSTGRES":
                    cur.execute(
                        "SELECT 1 FROM pg_database WHERE datname = %s",
                        (dbname,))
                    if cur.fetchone() is not None:
                        _log_info(f"DB <{db_id}> esiste gia' (POSTGRES)")
                    else:
                        cur.execute(
                            f"CREATE DATABASE {_quote_ident(db_type, dbname)}")
                        conn.commit()
                        _log_info(
                            f"DB <{db_id}> creato ({db_type}, DB={dbname})")
                elif db_type == "SQLSERVER":
                    cur.execute(
                        "SELECT 1 FROM sys.databases WHERE name = ?",
                        (dbname,))
                    if cur.fetchone() is not None:
                        _log_info(f"DB <{db_id}> esiste gia' (SQLSERVER)")
                    else:
                        cur.execute(
                            "CREATE DATABASE "
                            + _quote_ident(db_type, dbname))
                        conn.commit()
                        _log_info(
                            f"DB <{db_id}> creato ({db_type}, DB={dbname})")
                elif db_type == "MARIADB":
                    cur.execute(
                        "CREATE DATABASE IF NOT EXISTS "
                        f"{_quote_ident(db_type, dbname)}")
                    conn.commit()
                    _log_info(
                        f"DB <{db_id}> creato ({db_type}, DB={dbname})")
                else:  # ORACLE: creazione service/utente, policy DBA
                    cur.execute(
                        f"CREATE DATABASE {_quote_ident(db_type, dbname)}")
                    conn.commit()
                    _log_info(
                        f"DB <{db_id}> creato ({db_type}, DB={dbname})")
            finally:
                cur.close()
                conn.close()
        else:
            return _error_proc(f"TYPE non supportato: {db_type}", sProc)
        cfg["HANDLE"] = None
        cfg["STATUS"] = "CLOSE"
        dictDb[db_id] = cfg
        _db_used = True
        sResult = _do_return(db_id)
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def cmd_db_close(dictJob: Dict[str, str]) -> str:
    """DB.CLOSE: DB.ID=<id>."""
    global _db_used, dictDb
    sProc = "cmd_db_close"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        sResult, cfg = _db_cfg(db_id, False)
        if sResult:
            return _error_proc(sResult, sProc)
        conn = cfg.get("HANDLE")
        if cfg.get("STATUS") != "OPEN":
            return _error_proc(f"DB <{db_id}> non aperto", sProc)
        if conn is not None:
            conn.close()
        cfg["HANDLE"] = None
        cfg["STATUS"] = "CLOSE"
        _db_used = True
        _log_info(f"DB <{db_id}> chiuso")
        sResult = _do_return(db_id)
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def cmd_sql_exec(dictJob: Dict[str, str]) -> str:
    """SQL.EXEC: DB.ID + SQL xor FILE.IN.SQL, opz. FILE.OUT.XLS."""
    global _db_used, dictDb
    sProc = "cmd_sql_exec"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        sResult, cfg = _db_cfg(db_id, True)
        if sResult:
            return _error_proc(sResult, sProc)
        sql = str(dictJob.get("SQL", "")).strip()
        sql_file = str(dictJob.get("FILE.IN.SQL", "")).strip()
        if sql and sql_file:
            return _error_proc(
                "Specificare SQL xor FILE.IN.SQL, non entrambi", sProc)
        if not sql and not sql_file:
            return _error_proc("SQL o FILE.IN.SQL obbligatori", sProc)
        if sql_file:
            if not os.path.isfile(sql_file):
                return _error_proc(
                    f"FILE.IN.SQL non esistente: {sql_file}", sProc)
            with open(sql_file, "r", encoding="utf-8") as f:
                sql = f.read().strip()
        if not sql.upper().split()[0:1]:
            return _error_proc("SQL vuoto", sProc)
        db_type = cfg.get("TYPE", "")
        conn = cfg["HANDLE"]
        cur = conn.cursor()
        cur.execute(sql)
        if sql.upper().startswith("SELECT"):
            rows = cur.fetchall()
            file_out = str(dictJob.get("FILE.OUT.XLS", "")).strip()
            if file_out:
                # Scrittura Excel (riuso logica xls_merge in ntj_fact)
                import pandas as pd
                cols = [d[0] for d in cur.description]
                df = pd.DataFrame(rows, columns=cols)
                s_ext = file_out.lower()
                if s_ext.endswith(".xlsx"):
                    df.to_excel(file_out, index=False)
                elif s_ext.endswith(".xls"):
                    import xlwt
                    wb = xlwt.Workbook()
                    ws = wb.add_sheet("Sheet1")
                    for c, h in enumerate(cols):
                        ws.write(0, c, h)
                    for r, row in enumerate(rows, start=1):
                        for c, v in enumerate(row):
                            ws.write(r, c, v if v is not None else "")
                    wb.save(file_out)
                else:
                    raise ValueError("Estensione output non .xls/.xlsx")
                _log_info(f"Scritto {file_out} ({len(rows)} righe)")
            cur.close()
            _db_used = True
            sResult = _do_return(str(len(rows)),
                                 {"XLS": file_out} if file_out else None)
        else:
            conn.commit()
            rowcount = cur.rowcount if cur.rowcount >= 0 else 0
            cur.close()
            _db_used = True
            _log_info(f"Eseguito SQL ({rowcount} righe)")
            sResult = _do_return(str(rowcount))
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def cmd_tab_zap(dictJob: Dict[str, str]) -> str:
    """TAB.ZAP: DB.ID + TABLE."""
    global _db_used, dictDb
    sProc = "cmd_tab_zap"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        sResult, cfg = _db_cfg(db_id, True)
        if sResult:
            return _error_proc(sResult, sProc)
        table = str(dictJob.get("TABLE", "")).strip()
        if not table:
            return _error_proc("TABLE non precisata", sProc)
        db_type = cfg.get("TYPE", "")
        conn = cfg["HANDLE"]
        q = f"DELETE FROM {_quote_ident(db_type, table)}"
        cur = conn.cursor()
        cur.execute(q)
        conn.commit()
        rowcount = cur.rowcount if cur.rowcount >= 0 else 0
        cur.close()
        _db_used = True
        _log_info(f"TAB.ZAP {table} ({rowcount} righe)")
        sResult = _do_return(str(rowcount))
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def _load_struct(struct_file: str
                 ) -> Tuple[str, str, List[Dict[str, Any]], List[str]]:
    """Legge+valida JSON struttura. Ritorna (errore, table, cols, pk_list)."""
    sProc = "_load_struct"
    try:
        if not struct_file:
            return ("FILE.IN.STRUCT non precisato", "", [], [])
        if not os.path.isfile(struct_file):
            return (f"FILE.IN.STRUCT non esistente: {struct_file}", "", [], [])
        with open(struct_file, "r", encoding="utf-8") as f:
            j = json.load(f)
        table = str(j.get("table", "")).strip()
        if not table:
            return ("table mancante in JSON", "", [], [])
        if not re.match(r"^[A-Za-z0-9_]+$", table) or len(table) > 64:
            return (f"table non valida in JSON: {table}", "", [], [])
        cols = j.get("columns", [])
        if not cols:
            return ("columns mancanti in JSON", "", [], [])
        for c in cols:
            nm = str(c.get("name", ""))
            if not re.match(r"^[A-Za-z0-9_]+$", nm) or len(nm) > 64:
                return (f"colonna non valida in JSON: {nm}", "", [], [])
        pk_list = j.get("pk", [])
        if not isinstance(pk_list, list):
            pk_list = []
        return ("", table, cols, pk_list)
    except Exception as e:
        return (f"{sProc}: Errore {e}", "", [], [])


def _table_columns_db(conn: Any, db_type: str,
                      table: str) -> Tuple[str, List[str]]:
    """Colonne esistenti via SELECT WHERE 1=0. Ritorna (errore, [nomi])."""
    try:
        cur = conn.cursor()
        cur.execute(
            f"SELECT * FROM {_quote_ident(db_type, table)} WHERE 1=0")
        names = [d[0] for d in (cur.description or [])]
        cur.close()
        if not names:
            return (f"tabella {table}: nessuna colonna rilevata", [])
        return ("", names)
    except Exception:
        return (f"tabella {table} non esistente o illeggibile", [])


def cmd_tab_create_json(dictJob: Dict[str, str]) -> str:
    """TAB.CREATE.JSON: DB.ID + FILE.IN.STRUCT [+ TABLE opz]."""
    global _db_used, dictDb
    sProc = "cmd_tab_create_json"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        sResult, cfg = _db_cfg(db_id, True)
        if sResult:
            return _error_proc(sResult, sProc)
        struct_file = str(dictJob.get("FILE.IN.STRUCT", "")).strip()
        table_req = str(dictJob.get("TABLE", "")).strip()
        sResult, table, cols, pk_list = _load_struct(struct_file)
        if sResult:
            return _error_proc(sResult, sProc)
        if table_req and table_req != table:
            return _error_proc(
                f"TABLE richiesto ({table_req}) != JSON ({table})", sProc)
        conn = cfg["HANDLE"]
        db_type = cfg.get("TYPE", "")
        if _tab_exists(conn, db_type, table):
            return _error_proc(f"Tabella {table} gia' esistente", sProc)
        parts = []
        for i, c in enumerate(cols):
            ddl = _type_to_ddl(db_type, c)
            cname = _quote_ident(db_type, c.get("name", f"C{i}"))
            # pk composta gestita a livello tabella
            if c.get("pk", False) and isinstance(pk_list, list) and len(pk_list) > 1:
                pass  # pk composta gestita dopo
            parts.append(f"{cname} {ddl}")
        if isinstance(pk_list, list) and len(pk_list) > 1:
            pk_cols = ", ".join(_quote_ident(db_type, x) for x in pk_list)
            parts.append(f"PRIMARY KEY ({pk_cols})")
        sql = (
            f"CREATE TABLE {_quote_ident(db_type, table)} (" +
            ", ".join(parts) + ")")
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        cur.close()
        _db_used = True
        _log_info(f"Tabella {table} creata")
        sResult = _do_return(table)
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def cmd_db_import(dictJob: Dict[str, str]) -> str:
    """DB.IMPORT: DB.ID + [FILE.IN.STRUCT] + FILE.IN.CSV xor FILE.IN.JSON.

    STRUCT facoltativa: se data, il file importato deve rispettarla;
    se omessa, TABLE e' obbligatorio e le colonne si leggono dalla
    tabella esistente (tutte nullable, valori pass-through).
    """
    global _db_used, dictDb
    sProc = "cmd_db_import"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        sResult, cfg = _db_cfg(db_id, True)
        if sResult:
            return _error_proc(sResult, sProc)
        struct_file = str(dictJob.get("FILE.IN.STRUCT", "")).strip()
        file_csv = str(dictJob.get("FILE.IN.CSV", "")).strip()
        file_json = str(dictJob.get("FILE.IN.JSON", "")).strip()
        table_req = str(dictJob.get("TABLE", "")).strip()
        mode = str(dictJob.get("MODE", "APPEND")).upper()
        if not file_csv and not file_json:
            return _error_proc("FILE.IN.CSV xor FILE.IN.JSON obbligatori", sProc)
        if file_csv and file_json:
            return _error_proc(
                "Specificare FILE.IN.CSV xor FILE.IN.JSON, non entrambi", sProc)
        if mode not in ("APPEND", "REPLACE"):
            return _error_proc(f"MODE non valido: {mode} (APPEND|REPLACE)", sProc)
        conn = cfg["HANDLE"]
        db_type = cfg.get("TYPE", "")
        if struct_file:
            sResult, table, cols, _pk = _load_struct(struct_file)
            if sResult:
                return _error_proc(sResult, sProc)
            if table_req and table_req != table:
                return _error_proc(
                    f"TABLE richiesto ({table_req}) != JSON ({table})", sProc)
            col_names = [c.get("name") for c in cols]
            col_nullable = {c.get("name"): c.get("null", True) for c in cols}
            col_auto = {c.get("name"): bool(c.get("auto", False)) for c in cols}
        else:
            if not table_req:
                return _error_proc(
                    "TABLE obbligatorio senza FILE.IN.STRUCT", sProc)
            table = table_req
            sResult, col_names = _table_columns_db(conn, db_type, table)
            if sResult:
                return _error_proc(sResult, sProc)
            col_nullable = {cn: True for cn in col_names}
            col_auto = {}
        if mode == "REPLACE":
            # TAB.ZAP preliminare
            q = f"DELETE FROM {_quote_ident(db_type, table)}"
            cur = conn.cursor()
            cur.execute(q)
            conn.commit()
            cur.close()
        nLoaded = 0
        # Caricamento dati
        if file_csv:
            if not os.path.isfile(file_csv):
                return _error_proc(f"FILE.IN.CSV non esistente: {file_csv}", sProc)
            with open(file_csv, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f, delimiter=";")
                headers = reader.fieldnames or []
                # Match case-insensitive
                map_src_dst = {}
                for h in headers:
                    for cn in col_names:
                        if h.lower() == cn.lower():
                            map_src_dst[h] = cn
                            break
                missing = [cn for cn in col_names if cn not in map_src_dst.values()]
                ins_cols = [cn for cn in col_names if cn in map_src_dst.values()]
                for mn in missing:
                    if col_auto.get(mn, False):
                        continue  # generata dal DB, non richiesta nel file
                    if not col_nullable.get(mn, True):
                        return _error_proc(
                            f"Colonna obbligatoria mancante: {mn}", sProc)
                placeholders = [_placeholder(db_type, i+1) for i in range(len(ins_cols))]
                ins_sql = (
                    f"INSERT INTO {_quote_ident(db_type, table)} "
                    f"({', '.join(_quote_ident(db_type, n) for n in ins_cols)}) "
                    f"VALUES ({', '.join(placeholders)})")
                cur = conn.cursor()
                batch = []
                for i, row in enumerate(reader, start=2):
                    vals = []
                    for cn in ins_cols:
                        src = None
                        for sh, dh in map_src_dst.items():
                            if dh == cn:
                                src = row.get(sh, "")
                                break
                        if src is None or src == "":
                            if not col_nullable.get(cn, True):
                                return _error_proc(
                                    f"riga {i}: campo {cn} vuoto e non nullable",
                                    sProc)
                            vals.append(None)
                        else:
                            vals.append(src)
                    try:
                        batch.append(vals)
                        nLoaded += 1
                        if len(batch) >= 1000:
                            cur.executemany(ins_sql, batch)
                            batch = []
                    except Exception as e:
                        return _error_proc(f"errore riga {i}: {e}", sProc)
                if batch:
                    cur.executemany(ins_sql, batch)
                conn.commit()
                cur.close()
                _log_info(f"Import CSV {file_csv} in {table} ({nLoaded} righe)")
        elif file_json:
            if not os.path.isfile(file_json):
                return _error_proc(f"FILE.IN.JSON non esistente: {file_json}", sProc)
            with open(file_json, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, list):
                return _error_proc("JSON dati deve essere array di oggetti", sProc)
            for i, obj in enumerate(data, start=1):
                if not isinstance(obj, dict):
                    return _error_proc(f"riga {i}: non e' un oggetto", sProc)
            union = set()
            for obj in data:
                for cn in col_names:
                    v = obj.get(cn)
                    if v is not None and not (isinstance(v, str) and v == ""):
                        union.add(cn)
            ins_cols = [cn for cn in col_names if cn in union]
            for cn in col_names:
                if cn not in union:
                    if col_auto.get(cn, False):
                        continue
                    if not col_nullable.get(cn, True):
                        return _error_proc(
                            f"Colonna obbligatoria mai valorizzata: {cn}", sProc)
            placeholders = [_placeholder(db_type, i+1) for i in range(len(ins_cols))]
            ins_sql = (
                f"INSERT INTO {_quote_ident(db_type, table)} "
                f"({', '.join(_quote_ident(db_type, n) for n in ins_cols)}) "
                f"VALUES ({', '.join(placeholders)})")
            cur = conn.cursor()
            batch = []
            for i, obj in enumerate(data, start=1):
                vals = []
                for cn in ins_cols:
                    v = obj.get(cn)
                    if v is None or (isinstance(v, str) and v == ""):
                        if not col_nullable.get(cn, True):
                            return _error_proc(
                                f"riga {i}: campo {cn} nullo e non nullable",
                                sProc)
                        vals.append(None)
                    else:
                        vals.append(v)
                try:
                    batch.append(vals)
                    nLoaded += 1
                    if len(batch) >= 1000:
                        cur.executemany(ins_sql, batch)
                        batch = []
                except Exception as e:
                    return _error_proc(f"errore riga {i}: {e}", sProc)
            if batch:
                cur.executemany(ins_sql, batch)
            conn.commit()
            cur.close()
            _log_info(f"Import JSON {file_json} in {table} ({nLoaded} righe)")
        _db_used = True
        sResult = _do_return(str(nLoaded))
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def _export_value(v: Any, for_json: bool) -> Any:
    """Converte valore DB per export CSV/JSON (None, date, bool)."""
    import datetime
    if v is None:
        return None if for_json else ""
    if isinstance(v, bool):
        return v if for_json else (1 if v else 0)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    return v if for_json else str(v)


def cmd_db_export(dictJob: Dict[str, str]) -> str:
    """DB.EXPORT: DB.ID + OUT.QUERY + FILE.OUT.QUERY + OUT.TYPE (CSV|JSON)."""
    global _db_used, dictDb
    sProc = "cmd_db_export"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        sResult, cfg = _db_cfg(db_id, True)
        if sResult:
            return _error_proc(sResult, sProc)
        sql = str(dictJob.get("OUT.QUERY", "")).strip()
        file_out = str(dictJob.get("FILE.OUT.QUERY", "")).strip()
        out_type = str(dictJob.get("OUT.TYPE", "")).strip().upper()
        if not sql:
            return _error_proc("OUT.QUERY non precisata", sProc)
        if sql.upper().split()[0:1] != ["SELECT"]:
            return _error_proc("OUT.QUERY deve essere una SELECT", sProc)
        if out_type not in ("CSV", "JSON"):
            return _error_proc(
                f"OUT.TYPE non valido: {out_type} (CSV|JSON)", sProc)
        if not file_out:
            return _error_proc("FILE.OUT.QUERY non precisato", sProc)
        s_ext = ".csv" if out_type == "CSV" else ".json"
        if not file_out.lower().endswith(s_ext):
            return _error_proc(
                f"FILE.OUT.QUERY deve finire con {s_ext} "
                f"(OUT.TYPE={out_type})", sProc)
        conn = cfg["HANDLE"]
        cur = conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        cols = [d[0] for d in (cur.description or [])]
        cur.close()
        if out_type == "CSV":
            with open(file_out, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(cols)
                for row in rows:
                    writer.writerow([_export_value(v, False) for v in row])
        else:
            data = []
            for row in rows:
                data.append({c: _export_value(v, True)
                             for c, v in zip(cols, row)})
            with open(file_out, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        _db_used = True
        _log_info(f"Export {out_type} {file_out} ({len(rows)} righe)")
        sResult = _do_return(str(len(rows)), {"QUERY": file_out})
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def _compact_sqlite(sPath: str, timeout: int) -> None:
    """VACUUM in place (errore se transazione aperta o file danneggiato)."""
    import sqlite3
    conn = sqlite3.connect(sPath, timeout=timeout)
    try:
        conn.execute("VACUUM")
        conn.commit()
    finally:
        conn.close()


def _compact_access(sPath: str) -> None:
    """Compact and Repair via DAO (ACE) con fallback JRO (solo .mdb)."""
    sLow = sPath.lower()
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "ACCESS COMPACT richiede pywin32 (pip install pywin32)")
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    sTmp = sPath + ".compact_tmp"
    if os.path.isfile(sTmp):
        os.remove(sTmp)
    try:
        try:
            engine = win32com.client.Dispatch("DAO.DBEngine.120")
        except Exception:
            if sLow.endswith(".accdb"):
                raise RuntimeError(
                    "ACCESS COMPACT .accdb richiede Access Database "
                    "Engine (DAO.DBEngine.120 assente)")
            engine = win32com.client.Dispatch("DAO.DBEngine.36")
        engine.CompactDatabase(sPath, sTmp)
    except Exception:
        if os.path.isfile(sTmp):
            os.remove(sTmp)
        raise
    os.replace(sTmp, sPath)


def cmd_db_compact(dictJob: Dict[str, str]) -> str:
    """DB.COMPACT: FILE.IN.DB=<file> (.db/.sqlite/.mdb/.accdb).

    Prima crea backup <file>.old (sovrascrive), poi compatta.
    Il file non deve essere aperto in dictDb (prima DB.CLOSE).
    Solo file-DB (SQLITE/ACCESS); server via SQL.EXEC (vedi prompt).
    Ritorna byte recuperati in RETURN.VALUE.
    """
    global _db_used, dictDb
    sProc = "cmd_db_compact"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        sPath = str(dictJob.get("FILE.IN.DB", "")).strip()
        if not sPath:
            return _error_proc("FILE.IN.DB non precisato", sProc)
        if not os.path.isfile(sPath):
            return _error_proc(f"FILE.IN.DB non esistente: {sPath}", sProc)
        sLow = sPath.lower()
        if sLow.endswith((".db", ".sqlite", ".sqlite3", ".db3")):
            sKind = "SQLITE"
        elif sLow.endswith(".mdb") or sLow.endswith(".accdb"):
            sKind = "ACCESS"
        else:
            return _error_proc(
                f"estensione non compattabile: {sPath} "
                "(solo .db/.sqlite/.mdb/.accdb)", sProc)
        sAbs = os.path.abspath(os.path.normcase(sPath))
        for sId, cfg in dictDb.items():
            if cfg.get("STATUS") != "OPEN":
                continue
            sDbPath = str(cfg.get("FILE.IN.DB", ""))
            if sDbPath and os.path.abspath(
                    os.path.normcase(sDbPath)) == sAbs:
                return _error_proc(
                    f"file aperto come DB <{sId}>: prima DB.CLOSE", sProc)
        nBefore = os.path.getsize(sPath)
        shutil.copyfile(sPath, sPath + ".old")
        sLastErr = ""
        for nTry in (1, 2, 3):
            try:
                if sKind == "SQLITE":
                    _compact_sqlite(sPath, 30)
                else:
                    _compact_access(sPath)
                sLastErr = ""
                break
            except Exception as e:
                sLastErr = str(e)
                if nTry < 3:
                    _log_info(f"Compact tentativo {nTry} fallito, "
                              f"riprovo ({sLastErr[:100]})")
                    time.sleep(2)
        if sLastErr:
            return _error_proc(sLastErr, sProc)
        nAfter = os.path.getsize(sPath)
        nSaved = nBefore - nAfter
        _db_used = True
        _log_info(f"Compattato {sPath} ({sKind}): "
                  f"{nBefore} -> {nAfter} byte")
        sResult = _do_return(str(nSaved))
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


# =============================================================================
# CLEANUP E DISPATCHER
# =============================================================================

def db_cleanup_all() -> str:
    """Chiude tutti i DB aperti (chiamata a fine run da ntj_fact.py)."""
    global dictDb
    sProc = "db_cleanup_all"
    sInfo = []
    try:
        for db_id, cfg in dictDb.items():
            conn = cfg.get("HANDLE")
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
                cfg["HANDLE"] = None
                cfg["STATUS"] = "CLOSE"
                sInfo.append(f"chiuso {db_id}")
        if sInfo:
            _log_info("DB cleanup: " + ", ".join(sInfo))
        return ""
    except Exception as e:
        return _error_proc(str(e), sProc)


def cbCommands_db(dictJob: Dict[str, str]) -> str:
    """Dispatcher comandi DB.* per ntj_fact.cbCommands."""
    global _db_used
    sCommand = str(dictJob.get("COMMAND", "")).strip().upper()
    if sCommand == "DB.OPEN":
        return cmd_db_open(dictJob)
    elif sCommand == "DB.CREATE":
        return cmd_db_create(dictJob)
    elif sCommand == "DB.CLOSE":
        return cmd_db_close(dictJob)
    elif sCommand == "SQL.EXEC":
        return cmd_sql_exec(dictJob)
    elif sCommand == "TAB.ZAP":
        return cmd_tab_zap(dictJob)
    elif sCommand == "TAB.CREATE.JSON":
        return cmd_tab_create_json(dictJob)
    elif sCommand == "DB.IMPORT":
        return cmd_db_import(dictJob)
    elif sCommand == "DB.EXPORT":
        return cmd_db_export(dictJob)
    elif sCommand == "DB.COMPACT":
        return cmd_db_compact(dictJob)
    else:
        return f"cbCommands_db: Errore COMMAND DB sconosciuto: {sCommand}"


def is_db_used() -> bool:
    """True se almeno un comando DB e' stato eseguito in questo run."""
    return _db_used
