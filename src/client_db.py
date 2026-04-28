import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

DB_PATH = Path("data/aurumcalc.db")


def init_db(db_path: Path = DB_PATH) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
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
        )
        conn.commit()


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def upsert_project(client_name: str, project_name: str, payload: Dict[str, Any], db_path: Path = DB_PATH) -> None:
    init_db(db_path)
    now = _utc_now_iso()
    payload_json = json.dumps(payload, ensure_ascii=False, default=str)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO client_projects (client_name, project_name, payload_json, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(client_name, project_name)
            DO UPDATE SET payload_json=excluded.payload_json, updated_at=excluded.updated_at
            """,
            (client_name.strip(), project_name.strip(), payload_json, now, now),
        )
        conn.commit()


def list_projects(client_name: Optional[str] = None, db_path: Path = DB_PATH) -> List[Tuple[str, str, str]]:
    init_db(db_path)
    with sqlite3.connect(db_path) as conn:
        if client_name and client_name.strip():
            rows = conn.execute(
                """
                SELECT client_name, project_name, updated_at
                FROM client_projects
                WHERE client_name = ?
                ORDER BY updated_at DESC
                """,
                (client_name.strip(),),
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT client_name, project_name, updated_at
                FROM client_projects
                ORDER BY updated_at DESC
                """
            ).fetchall()
    return rows


def load_project(client_name: str, project_name: str, db_path: Path = DB_PATH) -> Optional[Dict[str, Any]]:
    init_db(db_path)
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            """
            SELECT payload_json
            FROM client_projects
            WHERE client_name = ? AND project_name = ?
            """,
            (client_name.strip(), project_name.strip()),
        ).fetchone()
    if not row:
        return None
    return json.loads(row[0])
