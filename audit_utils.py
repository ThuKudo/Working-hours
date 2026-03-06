import json
import os
import sqlite3
from typing import Any


def default_operator() -> str:
    return os.environ.get("USERNAME") or os.environ.get("USER") or "Unknown User"


def ensure_schema(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS work_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_name TEXT NOT NULL,
            work_date TEXT NOT NULL,
            project_name TEXT NOT NULL,
            task_name TEXT NOT NULL,
            hours_spent REAL NOT NULL CHECK (hours_spent > 0),
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS change_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id INTEGER,
            action TEXT NOT NULL,
            changed_by TEXT NOT NULL,
            employee_name TEXT,
            work_date TEXT,
            project_name TEXT,
            task_name TEXT,
            hours_spent REAL,
            old_values TEXT,
            new_values TEXT,
            changed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.commit()


def snapshot_from_row(row: sqlite3.Row | None) -> dict[str, Any]:
    if row is None:
        return {}
    return {
        "employee_name": row["employee_name"],
        "work_date": row["work_date"],
        "project_name": row["project_name"],
        "task_name": row["task_name"],
        "hours_spent": float(row["hours_spent"]),
    }


def log_change(
    conn: sqlite3.Connection,
    *,
    entry_id: int | None,
    action: str,
    changed_by: str,
    old_values: dict[str, Any] | None = None,
    new_values: dict[str, Any] | None = None,
) -> None:
    current = new_values or old_values or {}
    conn.execute(
        """
        INSERT INTO change_history (
            entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent,
            old_values, new_values
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            entry_id,
            action,
            changed_by,
            current.get("employee_name"),
            current.get("work_date"),
            current.get("project_name"),
            current.get("task_name"),
            current.get("hours_spent"),
            json.dumps(old_values or {}, ensure_ascii=False),
            json.dumps(new_values or {}, ensure_ascii=False),
        ),
    )


def fetch_history(
    conn: sqlite3.Connection,
    *,
    month: str | None = None,
    project: str | None = None,
    limit: int = 100,
) -> list[sqlite3.Row]:
    query = """
        SELECT id, entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, changed_at
        FROM change_history
        WHERE 1 = 1
    """
    params: list[Any] = []
    if month:
        query += " AND substr(work_date, 1, 7) = ?"
        params.append(month)
    if project and project != "All Projects":
        query += " AND project_name = ?"
        params.append(project)
    query += " ORDER BY changed_at DESC, id DESC LIMIT ?"
    params.append(limit)
    return conn.execute(query, params).fetchall()
