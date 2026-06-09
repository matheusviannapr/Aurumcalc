"""
Persistência de projetos.
Backend: PostgreSQL (via DATABASE_URL) com fallback para SQLite local.

Para Streamlit Cloud, defina em st.secrets ou variável de ambiente:
  DATABASE_URL = "postgresql://user:password@host:5432/dbname"

Supabase (gratuito): https://supabase.com
  DATABASE_URL = "postgresql://postgres:<senha>@db.<ref>.supabase.co:5432/postgres"
"""
import json
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

DB_PATH = Path("data/aurumcalc.db")

# ---------------------------------------------------------------------------
# Detecção de backend
# ---------------------------------------------------------------------------
_DATABASE_URL: Optional[str] = None
_HAS_PSYCOPG2 = False

try:
    import psycopg2
    import psycopg2.extras
    _HAS_PSYCOPG2 = True
except ImportError:
    pass


def _get_database_url() -> Optional[str]:
    global _DATABASE_URL
    if _DATABASE_URL is not None:
        return _DATABASE_URL
    url = os.environ.get("DATABASE_URL", "")
    if not url:
        try:
            import streamlit as st
            url = st.secrets.get("DATABASE_URL", "")
        except Exception:
            pass
    _DATABASE_URL = url or None
    return _DATABASE_URL


def _use_postgres() -> bool:
    return _HAS_PSYCOPG2 and bool(_get_database_url())


# ---------------------------------------------------------------------------
# SQLite backend
# ---------------------------------------------------------------------------
_CREATE_SQL = """
CREATE TABLE IF NOT EXISTS client_projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_name TEXT NOT NULL,
    project_name TEXT NOT NULL,
    payload_json TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(client_name, project_name)
)
"""

_CREATE_SQL_PG = """
CREATE TABLE IF NOT EXISTS client_projects (
    id SERIAL PRIMARY KEY,
    client_name TEXT NOT NULL,
    project_name TEXT NOT NULL,
    payload_json TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(client_name, project_name)
)
"""


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def init_db() -> None:
    if _use_postgres():
        _pg_init()
    else:
        _sqlite_init()


# ---------------------------------------------------------------------------
# SQLite helpers
# ---------------------------------------------------------------------------
def _sqlite_init(db_path: Path = DB_PATH) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(_CREATE_SQL)
        conn.commit()


def _sqlite_upsert(client_name: str, project_name: str, payload_json: str, now: str) -> None:
    _sqlite_init()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO client_projects (client_name, project_name, payload_json, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(client_name, project_name)
            DO UPDATE SET payload_json=excluded.payload_json, updated_at=excluded.updated_at
            """,
            (client_name, project_name, payload_json, now, now),
        )
        conn.commit()


def _sqlite_list(client_name: Optional[str]) -> List[Tuple]:
    _sqlite_init()
    with sqlite3.connect(DB_PATH) as conn:
        if client_name:
            return conn.execute(
                "SELECT client_name, project_name, updated_at FROM client_projects WHERE client_name=? ORDER BY updated_at DESC",
                (client_name,),
            ).fetchall()
        return conn.execute(
            "SELECT client_name, project_name, updated_at FROM client_projects ORDER BY updated_at DESC"
        ).fetchall()


def _sqlite_load(client_name: str, project_name: str) -> Optional[str]:
    _sqlite_init()
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT payload_json FROM client_projects WHERE client_name=? AND project_name=?",
            (client_name, project_name),
        ).fetchone()
    return row[0] if row else None


def _sqlite_delete(client_name: str, project_name: str) -> None:
    _sqlite_init()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "DELETE FROM client_projects WHERE client_name=? AND project_name=?",
            (client_name, project_name),
        )
        conn.commit()


def _sqlite_export_all() -> List[Dict]:
    _sqlite_init()
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT client_name, project_name, payload_json, created_at, updated_at FROM client_projects ORDER BY updated_at DESC"
        ).fetchall()
    return [
        {"client_name": r[0], "project_name": r[1], "payload": json.loads(r[2]), "created_at": r[3], "updated_at": r[4]}
        for r in rows
    ]


def _sqlite_import_all(records: List[Dict]) -> int:
    _sqlite_init()
    count = 0
    now = _utc_now_iso()
    with sqlite3.connect(DB_PATH) as conn:
        for rec in records:
            try:
                payload_json = json.dumps(rec.get("payload", {}), ensure_ascii=False, default=str)
                conn.execute(
                    """
                    INSERT INTO client_projects (client_name, project_name, payload_json, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(client_name, project_name)
                    DO UPDATE SET payload_json=excluded.payload_json, updated_at=excluded.updated_at
                    """,
                    (rec["client_name"], rec["project_name"], payload_json, rec.get("created_at", now), rec.get("updated_at", now)),
                )
                count += 1
            except Exception:
                pass
        conn.commit()
    return count


# ---------------------------------------------------------------------------
# PostgreSQL helpers
# ---------------------------------------------------------------------------
def _pg_conn():
    url = _get_database_url()
    return psycopg2.connect(url)


def _pg_init() -> None:
    with _pg_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(_CREATE_SQL_PG)
        conn.commit()


def _pg_upsert(client_name: str, project_name: str, payload_json: str, now: str) -> None:
    _pg_init()
    with _pg_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO client_projects (client_name, project_name, payload_json, created_at, updated_at)
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT (client_name, project_name)
                DO UPDATE SET payload_json=EXCLUDED.payload_json, updated_at=EXCLUDED.updated_at
                """,
                (client_name, project_name, payload_json, now, now),
            )
        conn.commit()


def _pg_list(client_name: Optional[str]) -> List[Tuple]:
    _pg_init()
    with _pg_conn() as conn:
        with conn.cursor() as cur:
            if client_name:
                cur.execute(
                    "SELECT client_name, project_name, updated_at FROM client_projects WHERE client_name=%s ORDER BY updated_at DESC",
                    (client_name,),
                )
            else:
                cur.execute("SELECT client_name, project_name, updated_at FROM client_projects ORDER BY updated_at DESC")
            return cur.fetchall()


def _pg_load(client_name: str, project_name: str) -> Optional[str]:
    _pg_init()
    with _pg_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT payload_json FROM client_projects WHERE client_name=%s AND project_name=%s",
                (client_name, project_name),
            )
            row = cur.fetchone()
    return row[0] if row else None


def _pg_delete(client_name: str, project_name: str) -> None:
    _pg_init()
    with _pg_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM client_projects WHERE client_name=%s AND project_name=%s",
                (client_name, project_name),
            )
        conn.commit()


def _pg_export_all() -> List[Dict]:
    _pg_init()
    with _pg_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT client_name, project_name, payload_json, created_at, updated_at FROM client_projects ORDER BY updated_at DESC"
            )
            rows = cur.fetchall()
    return [
        {"client_name": r[0], "project_name": r[1], "payload": json.loads(r[2]), "created_at": r[3], "updated_at": r[4]}
        for r in rows
    ]


def _pg_import_all(records: List[Dict]) -> int:
    _pg_init()
    count = 0
    now = _utc_now_iso()
    with _pg_conn() as conn:
        with conn.cursor() as cur:
            for rec in records:
                try:
                    payload_json = json.dumps(rec.get("payload", {}), ensure_ascii=False, default=str)
                    cur.execute(
                        """
                        INSERT INTO client_projects (client_name, project_name, payload_json, created_at, updated_at)
                        VALUES (%s, %s, %s, %s, %s)
                        ON CONFLICT (client_name, project_name)
                        DO UPDATE SET payload_json=EXCLUDED.payload_json, updated_at=EXCLUDED.updated_at
                        """,
                        (rec["client_name"], rec["project_name"], payload_json, rec.get("created_at", now), rec.get("updated_at", now)),
                    )
                    count += 1
                except Exception:
                    pass
        conn.commit()
    return count


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------
def upsert_project(client_name: str, project_name: str, payload: Dict[str, Any]) -> None:
    now = _utc_now_iso()
    payload_json = json.dumps(payload, ensure_ascii=False, default=str)
    if _use_postgres():
        _pg_upsert(client_name.strip(), project_name.strip(), payload_json, now)
    else:
        _sqlite_upsert(client_name.strip(), project_name.strip(), payload_json, now)


def list_projects(client_name: Optional[str] = None) -> List[Tuple[str, str, str]]:
    name = client_name.strip() if client_name and client_name.strip() else None
    if _use_postgres():
        return _pg_list(name)
    return _sqlite_list(name)


def load_project(client_name: str, project_name: str) -> Optional[Dict[str, Any]]:
    if _use_postgres():
        raw = _pg_load(client_name.strip(), project_name.strip())
    else:
        raw = _sqlite_load(client_name.strip(), project_name.strip())
    return json.loads(raw) if raw else None


def delete_project(client_name: str, project_name: str) -> None:
    if _use_postgres():
        _pg_delete(client_name.strip(), project_name.strip())
    else:
        _sqlite_delete(client_name.strip(), project_name.strip())


def export_all_projects() -> bytes:
    """Exporta todos os projetos como JSON bytes para download."""
    if _use_postgres():
        records = _pg_export_all()
    else:
        records = _sqlite_export_all()
    return json.dumps(records, ensure_ascii=False, indent=2, default=str).encode("utf-8")


def import_projects_from_json(data: bytes) -> int:
    """Importa projetos de um JSON exportado anteriormente. Retorna quantidade importada."""
    records = json.loads(data.decode("utf-8"))
    if not isinstance(records, list):
        return 0
    if _use_postgres():
        return _pg_import_all(records)
    return _sqlite_import_all(records)


def storage_backend() -> str:
    return "PostgreSQL" if _use_postgres() else "SQLite (local)"
