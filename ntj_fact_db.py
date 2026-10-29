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


def Timestamp() -> str:
    """Timestamp corrente nel formato ntJobsApp: YYYYMMDD:HHMMSS."""
    import datetime
    return datetime.datetime.now().strftime("%Y%m%d:%H%M%S")

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


def cmd_sql_exec(dictJob) -> str:
    """SQL.EXEC: DB.ID + SQL xor FILE.IN.SQL.

    SELECT: output opzionale via FILE.OUT.XLS (.xls/.xlsx) o
    FILE.OUT.EXPORT (.csv/.xls/.xlsx, per estensione), in xor.
    """
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
        file_xls = str(dictJob.get("FILE.OUT.XLS", "")).strip()
        file_exp = str(dictJob.get("FILE.OUT.EXPORT", "")).strip()
        if not sql.upper().startswith("SELECT") and (file_xls or file_exp):
            return _error_proc(
                "FILE.OUT.XLS/EXPORT previsti solo per SELECT", sProc)
        db_type = cfg.get("TYPE", "")
        conn = cfg["HANDLE"]
        cur = conn.cursor()
        cur.execute(sql)
        if sql.upper().startswith("SELECT"):
            rows = cur.fetchall()
            cols = [d[0] for d in (cur.description or [])]
            if file_xls and file_exp:
                return _error_proc(
                    "Specificare FILE.OUT.XLS xor FILE.OUT.EXPORT", sProc)
            file_out = ""
            sReturnKey = ""
            if file_xls:
                file_out, sReturnKey = file_xls, "XLS"
                if not file_out.lower().endswith((".xls", ".xlsx")):
                    return _error_proc(
                        "Estensione FILE.OUT.XLS non valida: "
                        "solo .xls/.xlsx", sProc)
            elif file_exp:
                file_out, sReturnKey = file_exp, "EXPORT"
                if not file_out.lower().endswith((".csv", ".xls", ".xlsx")):
                    return _error_proc(
                        "Estensione FILE.OUT.EXPORT non valida: "
                        "solo .csv/.xls/.xlsx", sProc)
            if file_out and file_out.lower().endswith(".csv"):
                # CSV stile DB.EXPORT: ;, header, utf-8-sig, NULL->vuoto
                with open(file_out, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f, delimiter=";")
                    writer.writerow(cols)
                    for row in rows:
                        writer.writerow([_export_value(v, False)
                                         for v in row])
            elif file_out and file_out.lower().endswith(".xlsx"):
                import pandas as pd
                df = pd.DataFrame(rows, columns=cols)
                df.to_excel(file_out, index=False)
            elif file_out:
                import xlwt
                wb = xlwt.Workbook()
                ws = wb.add_sheet("Sheet1")
                for c, h in enumerate(cols):
                    ws.write(0, c, h)
                for r, row in enumerate(rows, start=1):
                    for c, v in enumerate(row):
                        ws.write(r, c, v if v is not None else "")
                wb.save(file_out)
            if file_out:
                _log_info(f"Scritto {file_out} ({len(rows)} righe)")
            cur.close()
            _db_used = True
            sResult = _do_return(
                str(len(rows)),
                {sReturnKey: file_out} if file_out else None)
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
    """DB.EXPORT: DB.ID + OUT.QUERY + FILE.OUT.QUERY.

    Formato dal'estensione del file output: .csv/.json/.xlsx
    (altri tipi = errore).
    """
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
        if not sql:
            return _error_proc("OUT.QUERY non precisata", sProc)
        if sql.upper().split()[0:1] != ["SELECT"]:
            return _error_proc("OUT.QUERY deve essere una SELECT", sProc)
        if not file_out:
            return _error_proc("FILE.OUT.QUERY non precisato", sProc)
        f_low = file_out.lower()
        if f_low.endswith(".csv"):
            out_type = "CSV"
        elif f_low.endswith(".json"):
            out_type = "JSON"
        elif f_low.endswith(".xlsx"):
            out_type = "XLSX"
        else:
            s_ext = "." + file_out.rsplit(".", 1)[-1] if "." in \
                os.path.basename(file_out) else file_out
            return _error_proc(
                f"esportazione {s_ext} non supportata", sProc)
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
        elif out_type == "JSON":
            data = []
            for row in rows:
                data.append({c: _export_value(v, True)
                             for c, v in zip(cols, row)})
            with open(file_out, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        else:
            import pandas as pd
            df = pd.DataFrame(rows, columns=cols)
            df.to_excel(file_out, index=False)
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


def _info_dbms_version(conn: Any, db_type: str) -> str:
    """Versione DBMS per TYPE (best effort, 'n/d' se non disp.)."""
    try:
        cur = conn.cursor()
        if db_type == "SQLITE":
            import sqlite3
            cur.close()
            return f"SQLite {sqlite3.sqlite_version}"
        elif db_type == "SQLSERVER":
            cur.execute("SELECT @@VERSION")
            v = str(cur.fetchone()[0]).replace("\n", " ")
        elif db_type == "POSTGRES":
            cur.execute("SELECT version()")
            v = str(cur.fetchone()[0]).replace("\n", " ")
        elif db_type == "MARIADB":
            cur.execute("SELECT @@version")
            v = str(cur.fetchone()[0]).replace("\n", " ")
        elif db_type == "ORACLE":
            cur.execute(
                "SELECT banner FROM v$version WHERE ROWNUM=1")
            v = str(cur.fetchone()[0]).strip()
        else:  # ACCESS: versione driver via ODBC
            import pyodbc
            v = (f"driver {conn.getinfo(pyodbc.SQL_DRIVER_NAME)} "
                 f"{conn.getinfo(pyodbc.SQL_DRIVER_VER)}")
        cur.close()
        return v
    except Exception:
        return "n/d"


def _info_tables(conn: Any, db_type: str
                 ) -> Tuple[str, str, List[str]]:
    """Trova tabella count/colonne di count/list (errore, count, names)."""
    try:
        cur = conn.cursor()

        def close() -> None:
            cur.close()
    except Exception as e:
        return (str(e), "0", [])

    sCount = ""
    sNames = ""
    names: List[str] = []
    try:
        if db_type == "SQLITE":
            cur.execute(
                "SELECT name FROM sqlite_master "
                "WHERE type='table' AND name NOT LIKE 'sqlite_%'")
            names = [str(r[0]) for r in cur.fetchall()]
            sCount = "SELECT COUNT(*) FROM sqlite_master"
        elif db_type == "ORACLE":
            cur.execute(
                "SELECT table_name FROM user_tables")
            names = [str(r[0]) for r in cur.fetchall()]
            sCount = "SELECT COUNT(*) FROM user_tables"
        elif db_type in ("POSTGRES", "MARIADB", "SQLSERVER"):
            cur.execute(
                "SELECT table_name FROM information_schema.tables "
                "WHERE table_schema != 'INFORMATION_SCHEMA'")
            names = [str(r[0]) for r in cur.fetchall()]
            sCount = (
                "SELECT COUNT(*) FROM information_schema.tables "
                "WHERE table_schema != 'INFORMATION_SCHEMA'")
        elif db_type == "ACCESS":
            # MSysObjects di solito negata, best effort
            cur.execute(
                "SELECT Name FROM MSysObjects "
                "WHERE Type=1 AND Flags=0")
            names = [str(r[0]) for r in cur.fetchall()]
            sCount = (f"SELECT COUNT(*) FROM MSysObjects "
                      f"WHERE Type=1 AND Flags=0")
        else:
            return ("TYPE non gestito per elenco tabelle", "0", [])
        cur.close()
        if sCount:
            cur = conn.cursor()
            try:
                cur.execute(sCount)
                sNames = str(cur.fetchone()[0])
            finally:
                close()
    except Exception as e:
        return (str(e), "0", names)
    return ("", sNames, names)


def _link_source_connect(cfg: Dict[str, Any]) -> str:
    """Costruisce la Connect string per il link DAO dal TYPE sorgente.

    ACCESS sorgente: Jet-to-Jet (nessun driver extra). Altri TYPE:
    ODBC (la disponibilita' del driver dipende dall'ambiente).
    """
    t = str(cfg.get("TYPE", ""))
    if t == "ACCESS":
        return (";DATABASE="
                + os.path.abspath(str(cfg.get("FILE.IN.DB", ""))))
    user = str(cfg.get("USER", ""))
    pwd = str(cfg.get("PASSWORD", ""))
    host = str(cfg.get("HOST", "localhost"))
    port = str(cfg.get("PORT", ""))
    dbname = str(cfg.get("DBNAME", ""))
    if t == "SQLSERVER":
        driver = str(cfg.get("DRIVER", "ODBC Driver 17 for SQL Server"))
        if not port:
            port = "1433"
        cs = (f"ODBC;DRIVER={{{driver}}};SERVER={host},{port};"
              f"DATABASE={dbname};")
        if user:
            cs += f"UID={user};PWD={pwd};"
        else:
            cs += "Trusted_Connection=yes;"
        return cs
    if t == "SQLITE":
        driver = str(cfg.get("DRIVER", "SQLite3 ODBC Driver"))
        return ("ODBC;DRIVER={" + driver + "};DATABASE="
                + os.path.abspath(str(cfg.get("FILE.IN.DB", ""))) + ";")
    if t == "ORACLE":
        driver = str(cfg.get("DRIVER", ""))
        if not driver:
            raise ValueError(
                "driver ODBC per ORACLE non precisato (DRIVER)")
        service = str(cfg.get("SERVICE", ""))
        tns = str(cfg.get("TNSNAME", ""))
        target = tns or f"{host}:{port or '1521'}/{service}"
        return (f"ODBC;DRIVER={{{driver}}};DBQ={target};"
                f"UID={user};PWD={pwd};")
    if t in ("POSTGRES", "MARIADB"):
        driver = str(cfg.get("DRIVER", ""))
        if not driver:
            driver = ("PostgreSQL Unicode" if t == "POSTGRES"
                      else "MySQL ODBC 8.0 ANSI Driver")
        if t == "POSTGRES":
            if not port:
                port = "5432"
            cs = (f"ODBC;DRIVER={{{driver}}};SERVER={host};PORT={port};"
                  f"DATABASE={dbname};")
        else:
            if not port:
                port = "3306"
            cs = (f"ODBC;DRIVER={{{driver}}};SERVER={host};PORT={port};"
                  f"DATABASE={dbname};")
        cs += f"UID={user};PWD={pwd};"
        return cs
    raise ValueError(f"TYPE sorgente non gestito per DB.LINK: {t}")


def _acc_tab_exists(t_path: str, name: str) -> bool:
    """Probe esistenza tabella in file ACCESS via connessione FRESCA.

    L'handle ODBC gia' aperto sullo stesso file non vede le modifiche
    di schema fatte da altri handle (cache di catalogo ODBC): per
    verifiche su link DAO serve sempre una nuova connessione.
    """
    try:
        import pyodbc
        pyodbc.pooling = False
        cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
              f"DBQ={os.path.abspath(t_path)};")
        c = pyodbc.connect(cs, autocommit=True)
        try:
            return _tab_exists(c, "ACCESS", name)
        finally:
            c.close()
    except Exception:
        return False


def cmd_db_link(dictJob: Dict[str, str]) -> str:
    """DB.LINK: collega una tabella esterna in ACCESS come allegata.

    DB.ID = target ACCESS aperto; DBL.ID = sorgente aperta (altro ID);
    DBL.TABLE = tabella sorgente; DBL.NAME = nome della tabella
    allegata da creare nel target. Creazione via DAO CreateTableDef.
    """
    global _db_used, dictDb
    sProc = "cmd_db_link"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        sResult, target = _db_cfg(db_id, True)
        if sResult:
            return _error_proc(sResult, sProc)
        if str(target.get("TYPE", "")) != "ACCESS":
            return _error_proc(
                f"DB.ID <{db_id}> deve essere di tipo ACCESS "
                f"({target.get('TYPE')})", sProc)
        dbl_id = str(dictJob.get("DBL.ID", "")).strip().upper()
        if not dbl_id:
            return _error_proc("DBL.ID non precisato", sProc)
        if dbl_id == db_id:
            return _error_proc("DBL.ID = DB.ID: self-link non ammesso",
                               sProc)
        sResult, source = _db_cfg(dbl_id, True)
        if sResult:
            return _error_proc(sResult, sProc)
        dbl_table = str(dictJob.get("DBL.TABLE", "")).strip()
        if not dbl_table:
            return _error_proc("DBL.TABLE non precisata", sProc)
        dbl_name = str(dictJob.get("DBL.NAME", "")).strip()
        if not dbl_name:
            return _error_proc("DBL.NAME non precisato", sProc)
        if not re.match(r"^[A-Za-z0-9_]+$", dbl_name):
            return _error_proc(f"DBL.NAME non valido: {dbl_name}", sProc)
        src_type = str(source.get("TYPE", ""))
        s_conn = source.get("HANDLE")
        if not _tab_exists(s_conn, src_type, dbl_table):
            return _error_proc(
                f"tabella <{dbl_table}> non trovata nel sorgente "
                f"<{dbl_id}> ({src_type})", sProc)
        t_path = os.path.abspath(str(target.get("FILE.IN.DB", "")))
        if _acc_tab_exists(t_path, dbl_name):
            return _error_proc(
                f"tabella <{dbl_name}> gia' esistente nel target "
                f"<{db_id}>", sProc)
        try:
            import win32com.client
        except ImportError:
            return _error_proc(
                "DB.LINK richiede pywin32 (pip install pywin32)", sProc)
        connect_src = _link_source_connect(source)
        engine = win32com.client.Dispatch("DAO.DBEngine.120")
        dao_db = engine.OpenDatabase(t_path, False, False, "")
        try:
            td = dao_db.CreateTableDef(dbl_name)
            td.Connect = connect_src
            td.SourceTableName = dbl_table
            dao_db.TableDefs.Append(td)
        finally:
            dao_db.Close()
            del dao_db
        # Verifica post-link su connessione FRESCA al file (la stessa
        # regola del pre-check: l'handle aperto non vede lo schema nuovo)
        if not _acc_tab_exists(t_path, dbl_name):
            return _error_proc(
                f"link creato ma tabella <{dbl_name}> non visibile "
                "(riaprire DB.ID per vederla)", sProc)
        _db_used = True
        _log_info(f"DB.LINK {dbl_name} <- {dbl_id}.{dbl_table} "
                  f"({src_type})")
        sResult = _do_return(dbl_name)
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def cmd_db_info(dictJob: Dict[str, str]) -> str:
    """DB.INFO: FILE.OUT.INFO (.txt) con info DB opz. (DB.ID) + driver ODBC.

    DB.ID opzionale (se presente deve essere aperto): sezione DB con
    versione DBMS, parametri (PASSWORD nascosta), tabelle con COUNT e
    colonne. Senza DB.ID: solo ambiente + driver ODBC installati.
    Ogni sezione e' resiliente: n/d su errore.
    """
    global _db_used, dictDb
    sProc = "cmd_db_info"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        file_out = str(dictJob.get("FILE.OUT.INFO", "")).strip()
        if not file_out:
            return _error_proc("FILE.OUT.INFO non precisato", sProc)
        if not file_out.lower().endswith(".txt"):
            return _error_proc("FILE.OUT.INFO deve finire con .txt", sProc)
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        db_type = ""
        conn: Any = None
        cfg: Optional[Dict[str, Any]] = None
        if db_id:
            sResult, cfg = _db_cfg(db_id, True)
            if sResult:
                return _error_proc(sResult, sProc)
            db_type = str(cfg.get("TYPE", ""))
            conn = cfg.get("HANDLE")
        lines: List[str] = []
        import datetime
        import platform
        lines.append("DB.INFO - " + Timestamp())
        lines.append("")
        lines.append("AMBIENTE")
        lines.append("  PIATTFORMA: "
                     f"{sys.platform} / {platform.system()} "
                     f"({platform.machine()})")
        lines.append(f"  PYTHON: {platform.python_version()} "
                     f"({sys.executable})")
        lines.append("")

        try:
            import pyodbc
            lines.append("DRIVER ODBC INSTALLATI")
            for drv in pyodbc.drivers():
                lines.append(f"  {drv}")
            lines.append("")
        except ImportError:
            lines.append("DRIVER ODBC INSTALLATI")
            lines.append("  n/d (pyodbc non installato)")
            lines.append("")

        if conn is not None:
            lines.append(f"DB <{db_id}> ({db_type})")
            lines.append(f"  DBMS: {_info_dbms_version(conn, db_type)}")
            for k in ("FILE.IN.DB", "DBNAME", "HOST", "PORT", "SERVICE",
                      "TNSNAME", "USER", "SCHEMA", "CHARSET", "SSLMODE",
                      "READONLY", "TIMEOUT"):
                v = cfg.get(k, "")
                if v != "":
                    lines.append(f"  {k}: {v}")
            sErr, sNames, names = _info_tables(conn, db_type)
            lines.append(f"  TABELLE: {sNames}"
                         + (f" (elenco n/d: {sErr[:80]})" if sErr else ""))
            for name in names:
                try:
                    tcur = conn.cursor()
                    tcur.execute(
                        f"SELECT COUNT(*) FROM {_quote_ident(db_type, name)}")
                    n = str(tcur.fetchone()[0])
                    tcur.execute(
                        f"SELECT * FROM {_quote_ident(db_type, name)} "
                        "WHERE 1=0")
                    cols = ", ".join(
                        f"{d[0]}"
                        + (f"({d[1]})" if isinstance(d[1], str) and d[1]
                           else "")
                        for d in (tcur.description or []))
                    tcur.close()
                    lines.append(f"    {name}: {n} righe, "
                                 f"colonne: {cols}")
                except Exception as e:
                    try:
                        conn.rollback()
                    except Exception:
                        pass
                    lines.append(
                        f"    {name}: n/d ({str(e)[:80]})")
            lines.append("")
        with open(file_out, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        _db_used = True
        _log_info(f"DB.INFO -> {file_out} ({len(lines)} righe)")
        sResult = _do_return(str(len(lines)), {"INFO": file_out})
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


def _list_q(cur: Any, sql: str) -> List[List[str]]:
    """Esegue query catalogo e ritorna righe [[*], ...] (best effort)."""
    cur.execute(sql)
    return [[str(v) for v in row] for row in cur.fetchall()]


def _list_tables(conn: Any, db_type: str) -> Tuple[str, List[List[str]]]:
    """Nome tabelle per TYPE (senza tabelle di sistema)."""
    try:
        cur = conn.cursor()
        if db_type == "SQLITE":
            rows = _list_q(cur, "SELECT name FROM sqlite_master "
                                "WHERE type='table' AND name "
                                "NOT LIKE 'sqlite_%'")
        elif db_type == "SQLSERVER":
            rows = _list_q(cur, "SELECT table_name FROM "
                                "information_schema.tables WHERE "
                                "table_type='BASE TABLE'")
        elif db_type == "ORACLE":
            rows = _list_q(cur, "SELECT table_name FROM user_tables")
        elif db_type == "POSTGRES":
            rows = _list_q(cur, "SELECT table_name FROM "
                                "information_schema.tables WHERE "
                                "table_type='BASE TABLE' AND "
                                "table_schema NOT IN ('pg_catalog',"
                                "'information_schema')")
        elif db_type == "MARIADB":
            rows = _list_q(cur, "SELECT table_name FROM "
                                "information_schema.tables WHERE "
                                "table_type='BASE TABLE' AND "
                                "table_schema=DATABASE()")
        elif db_type == "ACCESS":
            rows = []
            for t in cur.tables(tableType="TABLE"):
                rows.append([str(t.table_name)])
        else:
            cur.close()
            return ("TYPE non gestito", [])
        cur.close()
        return ("", rows)
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return ("n/d", [])


def _list_views(conn: Any, db_type: str) -> Tuple[str, List[List[str]]]:
    """Viste per TYPE (ACCESS: n/a, le query salvate sono in QUERIES)."""
    try:
        if db_type == "ACCESS":
            return ("", [])
        cur = conn.cursor()
        if db_type == "SQLITE":
            rows = _list_q(cur, "SELECT name FROM sqlite_master "
                                "WHERE type='view'")
        elif db_type == "SQLSERVER":
            rows = _list_q(cur, "SELECT table_name FROM "
                                "information_schema.views")
        elif db_type == "ORACLE":
            rows = _list_q(cur, "SELECT view_name FROM user_views")
        elif db_type == "POSTGRES":
            rows = _list_q(cur, "SELECT table_name FROM "
                                "information_schema.tables WHERE "
                                "table_type='VIEW' AND table_schema "
                                "NOT IN ('pg_catalog',"
                                "'information_schema')")
        elif db_type == "MARIADB":
            rows = _list_q(cur, "SELECT table_name FROM "
                                "information_schema.tables WHERE "
                                "table_type='VIEW' AND "
                                "table_schema=DATABASE()")
        else:
            cur.close()
            return ("TYPE non gestito", [])
        cur.close()
        return ("", rows)
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return ("n/d", [])


def _list_queries_access(conn: Any, cfg: Dict[str, Any]
                         ) -> Tuple[str, List[List[str]]]:
    """Query salvate ACCESS via ADOX (Views=SELECT, Procedures=action)."""
    try:
        import win32com.client
    except ImportError:
        return ("pywin32 non installato", [])
    sProc = "_list_queries_access"
    cat = None
    try:
        path = os.path.abspath(str(cfg.get("FILE.IN.DB", "")))
        cat = win32com.client.Dispatch("ADOX.Catalog")
        cat.ActiveConnection = (
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";")
        rows: List[List[str]] = []
        for v in cat.Views:
            rows.append(["VIEW-QUERY", str(v.Name)])
        for p in cat.Procedures:
            rows.append(["ACTION-QUERY", str(p.Name)])
        return ("", rows)
    except Exception as e:
        return (str(e)[:120], [])
    finally:
        if cat is not None:
            del cat
        import gc
        gc.collect()


def _list_queries(conn: Any, db_type: str, cfg: Dict[str, Any]
                  ) -> Tuple[str, List[List[str]]]:
    """Query salvate (solo ACCESS via ADOX; altri TYPE n/a)."""
    if db_type != "ACCESS":
        return ("non previsto per " + db_type, [])
    return _list_queries_access(conn, cfg)


def _list_procs(conn: Any, db_type: str) -> Tuple[str, List[List[str]]]:
    """Store procedures/functions (ACCESS: n/a, le query sono in QUERIES)."""
    try:
        if db_type in ("ACCESS", "SQLITE"):
            return ("", [])
        cur = conn.cursor()
        if db_type == "SQLSERVER":
            rows = _list_q(cur, "SELECT routine_name, routine_type FROM "
                                "information_schema.routines")
        elif db_type == "ORACLE":
            rows = _list_q(cur, "SELECT object_name, object_type FROM "
                                "user_objects WHERE object_type IN "
                                "('PROCEDURE','FUNCTION','PACKAGE')")
        elif db_type == "POSTGRES":
            rows = _list_q(cur, "SELECT proname, CASE prokind WHEN 'p' "
                                "THEN 'PROCEDURE' ELSE 'FUNCTION' END "
                                "FROM pg_proc p JOIN pg_namespace n ON "
                                "p.pronamespace=n.oid WHERE n.nspname "
                                "NOT IN ('pg_catalog',"
                                "'information_schema')")
        elif db_type == "MARIADB":
            rows = _list_q(cur, "SELECT routine_name, routine_type FROM "
                                "information_schema.routines WHERE "
                                "routine_schema=DATABASE()")
        else:
            cur.close()
            return ("TYPE non gestito", [])
        cur.close()
        return ("", rows)
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return ("n/d", [])


def _list_triggers(conn: Any,
                   db_type: str) -> Tuple[str, List[List[str]]]:
    """Trigger per TYPE (info = tabella di aggancio se disponibile)."""
    try:
        if db_type == "ACCESS":
            return ("", [])
        cur = conn.cursor()
        if db_type == "SQLITE":
            rows = _list_q(cur, "SELECT name FROM sqlite_master "
                                "WHERE type='trigger'")
        elif db_type == "SQLSERVER":
            rows = _list_q(cur, "SELECT name FROM sys.triggers")
        elif db_type == "ORACLE":
            rows = _list_q(
                cur, "SELECT trigger_name, table_name FROM user_triggers")
        elif db_type == "POSTGRES":
            rows = _list_q(cur, "SELECT c.relname, t.tgname FROM pg_trigger "
                                "t JOIN pg_class c ON t.tgrelid=c.oid "
                                "WHERE NOT t.tgisinternal")
        elif db_type == "MARIADB":
            rows = _list_q(cur, "SELECT trigger_name, event_object_table "
                                "FROM information_schema.triggers WHERE "
                                "trigger_schema=DATABASE()")
        else:
            cur.close()
            return ("TYPE non gestito", [])
        cur.close()
        return ("", rows)
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return ("n/d", [])


def _list_indexes(conn: Any,
                  db_type: str) -> Tuple[str, List[List[str]]]:
    """Indici per TYPE (info = tabella + unique se disponibile)."""
    try:
        cur = conn.cursor()
        rows: List[List[str]] = []
        if db_type == "SQLITE":
            cur.execute("SELECT name FROM sqlite_master WHERE "
                        "type='table' AND name NOT LIKE 'sqlite_%'")
            tables = [str(r[0]) for r in cur.fetchall()]
            for t in tables:
                cur.execute(f"PRAGMA index_list({_quote_ident(db_type, t)})")
                for r in cur.fetchall():
                    rows.append([str(r[1]), f"{t} unique={r[2]}"])
        elif db_type in ("ACCESS", "SQLSERVER"):
            for t in _list_tables(conn, db_type)[1]:
                try:
                    for r in cur.indexes(table=t[0]):
                        rows.append([str(r[2]),
                                     f"{t[0]} unique={bool(r[4])}"])
                except Exception:
                    pass
        elif db_type == "ORACLE":
            rows = _list_q(cur, "SELECT index_name, table_name FROM "
                                "user_indexes")
        elif db_type == "POSTGRES":
            rows = _list_q(cur, "SELECT indexname, tablename FROM "
                                "pg_indexes WHERE schemaname NOT IN "
                                "('pg_catalog','information_schema')")
        elif db_type == "MARIADB":
            rows = _list_q(cur, "SELECT DISTINCT index_name, table_name "
                                "FROM information_schema.statistics WHERE "
                                "index_name IS NOT NULL AND index_name "
                                "<> 'PRIMARY' AND "
                                "table_schema=DATABASE()")
        else:
            cur.close()
            return ("TYPE non gestito", [])
        cur.close()
        return ("", rows)
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return ("n/d", [])


def _list_pkfk(conn: Any, db_type: str) -> Tuple[str, List[List[str]]]:
    """PK/FK per TYPE (info = vincolo + colonne)."""
    try:
        cur = conn.cursor()
        rows: List[List[str]] = []
        if db_type == "SQLITE":
            cur.execute("SELECT name FROM sqlite_master WHERE "
                        "type='table' AND name NOT LIKE 'sqlite_%'")
            tables = [str(r[0]) for r in cur.fetchall()]
            for t in tables:
                cur.execute(f"PRAGMA table_info({_quote_ident(db_type, t)})")
                pkt = [str(r[1]) for r in cur.fetchall() if r[5]]
                if pkt:
                    rows.append([f"PK {t}", "(" + ",".join(pkt) + ")"])
                cur.execute(f"PRAGMA foreign_key_list({_quote_ident(db_type, t)})")
                for r in cur.fetchall():
                    rows.append([f"FK {t}->{r[2]}",
                                 f"{r[3]}->{str(r[4] or '')}"])
        elif db_type in ("ACCESS", "SQLSERVER"):
            for t in _list_tables(conn, db_type)[1]:
                tn = t[0]
                try:
                    for r in cur.primaryKeys(tn):
                        rows.append([f"PK {tn}",
                                     f"{r[2]} ({r[3]})"])
                except Exception:
                    pass
                try:
                    for r in cur.foreignKeys(table=tn):
                        rows.append([f"FK {r[6]}->{r[2]}",
                                     f"{r[7]}->{r[3]}"])
                except Exception:
                    pass
        elif db_type == "ORACLE":
            rows = _list_q(cur, "SELECT constraint_name, table_name, "
                                "constraint_type FROM user_constraints "
                                "WHERE constraint_type IN ('P','F')")
        elif db_type in ("POSTGRES", "MARIADB"):
            rows = _list_q(cur, "SELECT constraint_name, table_name, "
                                "constraint_type FROM "
                                "information_schema.table_constraints "
                                "WHERE constraint_type IN ('PRIMARY "
                                "KEY','FOREIGN KEY')"
                                + (" AND table_schema=DATABASE()"
                                   if db_type == "MARIADB" else
                                   " AND table_schema NOT IN "
                                   "('pg_catalog','information_schema')"))
        else:
            cur.close()
            return ("TYPE non gestito", [])
        cur.close()
        return ("", rows)
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return ("n/d", [])


def _list_plus(conn: Any, db_type: str) -> Tuple[str, List[List[str]]]:
    """PLUS: sequences/synonyms/events (engine-wise; SQLITE/ACCESS n/a)."""
    try:
        if db_type in ("SQLITE", "ACCESS"):
            return ("n/a", [])
        cur = conn.cursor()
        rows: List[List[str]] = []
        if db_type == "ORACLE":
            for s in _list_q(cur, "SELECT sequence_name FROM user_sequences"):
                rows.append(["SEQUENCE", s[0]])
            for s in _list_q(cur, "SELECT synonym_name, table_name FROM "
                                  "user_synonyms"):
                rows.append(["SYNONYM", f"{s[0]}->{s[1]}"])
        elif db_type == "SQLSERVER":
            try:
                for r in _list_q(cur, "SELECT name FROM sys.sequences"):
                    rows.append(["SEQUENCE", r[0]])
            except Exception:
                pass
            try:
                for r in _list_q(cur, "SELECT name FROM sys.synonyms"):
                    rows.append(["SYNONYM", r[0]])
            except Exception:
                pass
        elif db_type == "POSTGRES":
            for r in _list_q(cur, "SELECT c.relname FROM pg_class c JOIN "
                                  "pg_namespace n ON c.relnamespace=n.oid "
                                  "WHERE c.relkind='S' AND n.nspname NOT "
                                  "IN ('pg_catalog','information_schema')"):
                rows.append(["SEQUENCE", r[0]])
        elif db_type == "MARIADB":
            try:
                for r in _list_q(cur, "SELECT event_name FROM "
                                      "information_schema.events WHERE "
                                      "event_schema=DATABASE()"):
                    rows.append(["EVENT", r[0]])
            except Exception:
                pass
        cur.close()
        return ("", rows)
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return ("n/d", [])


_LIST_KINDS = ("TABLES", "VIEWS", "QUERIES", "PROCS", "TRIGGERS",
               "INDEXES", "PKFK", "PLUS")


def _list_kind(conn: Any, db_type: str, kind: str, cfg: Dict[str, Any]
               ) -> Tuple[str, List[List[str]]]:
    """Lista una KIND per TYPE. Ritorna (errore/n/d or '', rows)."""
    if kind == "TABLES":
        return _list_tables(conn, db_type)
    elif kind == "VIEWS":
        return _list_views(conn, db_type)
    elif kind == "QUERIES":
        return _list_queries(conn, db_type, cfg)
    elif kind == "PROCS":
        return _list_procs(conn, db_type)
    elif kind == "TRIGGERS":
        return _list_triggers(conn, db_type)
    elif kind == "INDEXES":
        return _list_indexes(conn, db_type)
    elif kind == "PKFK":
        return _list_pkfk(conn, db_type)
    elif kind == "PLUS":
        return _list_plus(conn, db_type)
    return (f"LIST non valido: {kind}", [])


def cmd_db_list(dictJob: Dict[str, str]) -> str:
    """DB.LIST: DB.ID + LIST (default ALL) + FILE.OUT.LIST.

    Lista oggetti del DB (senza dump del codice): per estensione del
    file output .csv/.json/.xlsx; categorie per KIND (TABLES tabelle,
    VIEWS viste, QUERIES query salvate solo ACCESS, PROCS procedures/
    functions, TRIGGERS, INDEXES, PKFK vincoli PK/FK, PLUS sequences/
    synonyms/events). In ALL le KIND non applicabili al TYPE sono
    saltate. Righe: KIND,NAME,INFO (KIND vuote -> "(vuoto)").
    """
    global _db_used, dictDb
    sProc = "cmd_db_list"
    print("Sto eseguendo " + sProc)
    sResult = ""
    try:
        db_id = str(dictJob.get("DB.ID", "")).strip().upper()
        sResult, cfg = _db_cfg(db_id, True)
        if sResult:
            return _error_proc(sResult, sProc)
        file_out = str(dictJob.get("FILE.OUT.LIST", "")).strip()
        if not file_out:
            return _error_proc("FILE.OUT.LIST non precisato", sProc)
        f_low = file_out.lower()
        if f_low.endswith(".csv"):
            out_type = "CSV"
        elif f_low.endswith(".json"):
            out_type = "JSON"
        elif f_low.endswith(".xlsx"):
            out_type = "XLSX"
        else:
            s_ext = "." + file_out.rsplit(".", 1)[-1] if "." in \
                os.path.basename(file_out) else file_out
            return _error_proc(
                f"esportazione {s_ext} non supportata", sProc)
        s_list = str(dictJob.get("LIST", "ALL")).strip().upper()
        kinds: List[str] = []
        if s_list == "ALL":
            kinds = list(_LIST_KINDS)
        elif s_list in _LIST_KINDS:
            kinds = [s_list]
        else:
            return _error_proc(
                f"LIST non valida: {s_list} "
                f"({','.join(_LIST_KINDS)}|ALL)", sProc)
        db_type = str(cfg.get("TYPE", ""))
        conn = cfg["HANDLE"]
        rows: List[List[str]] = []
        for kind in kinds:
            sErr, krows = _list_kind(conn, db_type, kind, cfg)
            if sErr and (sErr == "n/a"
                         or sErr.startswith("non previsto")):
                continue
            if sErr:
                rows.append([kind, "(n/d)", sErr[:120]])
                continue
            if not krows:
                rows.append([kind, "(vuoto)", ""])
                continue
            for r in krows:
                rows.append([kind, r[0], r[1] if len(r) > 1 else ""])
        if out_type == "CSV":
            with open(file_out, "w", encoding="utf-8-sig",
                      newline="") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(["KIND", "NAME", "INFO"])
                for r in rows:
                    writer.writerow(r)
        elif out_type == "JSON":
            data = [{"KIND": a, "NAME": b, "INFO": c}
                    for a, b, c in rows]
            with open(file_out, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        else:
            import pandas as pd
            df = pd.DataFrame(rows, columns=["KIND", "NAME", "INFO"])
            df.to_excel(file_out, index=False)
        _db_used = True
        _log_info(f"DB.LIST {s_list} {file_out} ({len(rows)} righe)")
        sResult = _do_return(str(len(rows)), {"LIST": file_out})
    except Exception as e:
        sResult = str(e)
    return _error_proc(sResult, sProc)


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
    elif sCommand == "DB.INFO":
        return cmd_db_info(dictJob)
    elif sCommand == "DB.LIST":
        return cmd_db_list(dictJob)
    elif sCommand == "DB.LINK":
        return cmd_db_link(dictJob)
    else:
        return f"cbCommands_db: Errore COMMAND DB sconosciuto: {sCommand}"


def is_db_used() -> bool:
    """True se almeno un comando DB e' stato eseguito in questo run."""
    return _db_used
