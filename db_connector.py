# =============================================================================
# db_connector.py – Direct SQL Server connection to MineStar database
# =============================================================================
from __future__ import annotations
import json
import logging
import os
from datetime import date
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)

CONN_SETTINGS_FILE = "db_settings.json"

# Try drivers in preference order
ODBC_DRIVERS = [
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server",
]


def get_available_driver() -> Optional[str]:
    """Return the best available SQL Server ODBC driver, or None."""
    try:
        import pyodbc
        installed = pyodbc.drivers()
        for preferred in ODBC_DRIVERS:
            if preferred in installed:
                return preferred
        # Fallback: any driver with "SQL Server" in the name
        fallback = [d for d in installed if "SQL Server" in d]
        return fallback[0] if fallback else None
    except ImportError:
        return None


def build_conn_str(
    server: str,
    database: str,
    use_windows_auth: bool,
    username: str = "",
    password: str = "",
    driver: Optional[str] = None,
) -> str:
    if driver is None:
        driver = get_available_driver() or "ODBC Driver 17 for SQL Server"
    base = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
    if use_windows_auth:
        return base + "Trusted_Connection=yes;"
    return base + f"UID={username};PWD={password};"


def test_connection(conn_str: str) -> tuple[bool, str]:
    """Return (success, message)."""
    try:
        import pyodbc
        conn = pyodbc.connect(conn_str, timeout=10)
        conn.close()
        return True, "Connected successfully."
    except ImportError:
        return False, (
            "pyodbc is not installed. Run:  pip install pyodbc\n"
            "Then install the ODBC driver from Microsoft if not already present."
        )
    except Exception as e:
        return False, str(e)


def get_tables(conn_str: str) -> list[str]:
    """Return all user tables as 'schema.table' strings."""
    import pyodbc
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT TABLE_SCHEMA + '.' + TABLE_NAME
        FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_TYPE = 'BASE TABLE'
        ORDER BY TABLE_SCHEMA, TABLE_NAME
    """)
    tables = [row[0] for row in cursor.fetchall()]
    conn.close()
    return tables


def get_date_columns(conn_str: str, table: str) -> list[str]:
    """Return datetime-typed columns in a table."""
    import pyodbc
    parts  = table.split(".", 1)
    schema = parts[0] if len(parts) == 2 else "dbo"
    tname  = parts[1] if len(parts) == 2 else parts[0]
    conn   = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?
          AND DATA_TYPE IN ('date','datetime','datetime2','smalldatetime')
        ORDER BY ORDINAL_POSITION
    """, schema, tname)
    cols = [row[0] for row in cursor.fetchall()]
    conn.close()
    return cols


def get_row_count(conn_str: str, table: str) -> int:
    try:
        import pyodbc
        conn   = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        count  = cursor.fetchone()[0]
        conn.close()
        return count
    except Exception:
        return -1


def preview_table(conn_str: str, table: str, n: int = 5) -> pd.DataFrame:
    import pyodbc
    conn = pyodbc.connect(conn_str)
    df   = pd.read_sql(f"SELECT TOP {n} * FROM {table}", conn)
    conn.close()
    return df


def score_table(conn_str: str, table: str) -> dict:
    """
    Quick-score a table for MineStar relevance.
    Returns {score, detected_types, row_count}.
    """
    KEYWORDS = {
        "cycle_times":  ["cycle", "queue", "travel", "haul"],
        "payload":      ["payload", "tonnes", "load weight"],
        "utilization":  ["availability", "utilization", "idle", "downtime", "operating hours"],
        "operators":    ["operator", "driver"],
    }
    TYPE_ICONS = {
        "cycle_times": "🔄 Cycle Times",
        "payload":     "⚖️ Payload",
        "utilization": "🔧 Utilization",
        "operators":   "👷 Operators",
    }
    try:
        df_preview = preview_table(conn_str, table, n=2)
        cols_lower = " ".join(df_preview.columns.str.lower())
        detected = [TYPE_ICONS[k] for k, kws in KEYWORDS.items()
                    if any(kw in cols_lower for kw in kws)]
        rc = get_row_count(conn_str, table)
        return {"score": len(detected), "detected_types": detected, "row_count": rc}
    except Exception:
        return {"score": 0, "detected_types": [], "row_count": -1}


def query_table(
    conn_str: str,
    table: str,
    date_col: Optional[str] = None,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
    limit: int = 50_000,
) -> pd.DataFrame:
    import pyodbc
    conn = pyodbc.connect(conn_str)
    where  = ""
    params = []
    if date_col and start_date and end_date:
        where  = f"WHERE [{date_col}] BETWEEN ? AND ?"
        params = [start_date, end_date]
    query = f"SELECT TOP {limit} * FROM {table} {where} ORDER BY (SELECT NULL)"
    df = pd.read_sql(query, conn, params=params if params else None)
    conn.close()
    return df


# ---------------------------------------------------------------------------
# Persist connection settings (no password saved)
# ---------------------------------------------------------------------------

def save_settings(server: str, database: str, use_windows_auth: bool,
                  username: str = "", last_table: str = "") -> None:
    try:
        with open(CONN_SETTINGS_FILE, "w") as f:
            json.dump({
                "server": server, "database": database,
                "use_windows_auth": use_windows_auth,
                "username": username, "last_table": last_table,
            }, f, indent=2)
    except Exception:
        pass


def load_settings() -> dict:
    if os.path.exists(CONN_SETTINGS_FILE):
        try:
            with open(CONN_SETTINGS_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {}
