import json
import os
import sqlite3
from pathlib import Path
from typing import Any

from audit_utils import ensure_schema


ALL_PROJECTS = "All Projects"
SUPABASE_DB_URL = os.environ.get("SUPABASE_DB_URL") or os.environ.get("DATABASE_URL")
USE_POSTGRES = bool(SUPABASE_DB_URL and SUPABASE_DB_URL.startswith("postgres"))

if USE_POSTGRES:
    import psycopg
    from psycopg.rows import dict_row


def init_backend(sqlite_db_path: Path) -> None:
    if USE_POSTGRES:
        _ensure_postgres_schema()
    else:
        with _sqlite_connection(sqlite_db_path) as conn:
            ensure_schema(conn)


def _sqlite_connection(sqlite_db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(sqlite_db_path)
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_postgres_schema() -> None:
    with psycopg.connect(SUPABASE_DB_URL) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS work_entries (
                    id BIGSERIAL PRIMARY KEY,
                    legacy_sqlite_id BIGINT,
                    employee_name TEXT NOT NULL,
                    work_date TEXT NOT NULL,
                    project_name TEXT NOT NULL,
                    task_name TEXT NOT NULL,
                    hours_spent NUMERIC(10,2) NOT NULL CHECK (hours_spent > 0),
                    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                );
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS change_history (
                    id BIGSERIAL PRIMARY KEY,
                    legacy_sqlite_history_id BIGINT,
                    entry_id BIGINT,
                    action TEXT NOT NULL,
                    changed_by TEXT NOT NULL,
                    employee_name TEXT,
                    work_date TEXT,
                    project_name TEXT,
                    task_name TEXT,
                    hours_spent NUMERIC(10,2),
                    old_values JSONB,
                    new_values JSONB,
                    changed_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                );
                """
            )
            cur.execute("ALTER TABLE work_entries ADD COLUMN IF NOT EXISTS legacy_sqlite_id BIGINT")
            cur.execute("ALTER TABLE change_history ADD COLUMN IF NOT EXISTS legacy_sqlite_history_id BIGINT")
            cur.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS uq_work_entries_legacy_sqlite_id ON work_entries (legacy_sqlite_id) WHERE legacy_sqlite_id IS NOT NULL;"
            )
            cur.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS uq_change_history_legacy_sqlite_id ON change_history (legacy_sqlite_history_id) WHERE legacy_sqlite_history_id IS NOT NULL;"
            )
            cur.execute("CREATE INDEX IF NOT EXISTS idx_work_entries_month ON work_entries (substring(work_date,1,7));")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_work_entries_project ON work_entries (project_name);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_change_history_changed_at ON change_history (changed_at DESC);")
        conn.commit()


def fetch_filters(sqlite_db_path: Path) -> dict:
    if USE_POSTGRES:
        with psycopg.connect(SUPABASE_DB_URL) as conn:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(
                    """
                    SELECT DISTINCT substring(work_date, 1, 7) AS month_key
                    FROM work_entries
                    ORDER BY month_key DESC
                    """
                )
                months = [row["month_key"] for row in cur.fetchall()]
                cur.execute("SELECT DISTINCT project_name FROM work_entries ORDER BY project_name")
                projects = [row["project_name"] for row in cur.fetchall()]
        return {"months": months, "projects": projects}

    with _sqlite_connection(sqlite_db_path) as conn:
        months = [
            row["month_key"]
            for row in conn.execute(
                """
                SELECT DISTINCT substr(work_date, 1, 7) AS month_key
                FROM work_entries
                ORDER BY month_key DESC
                """
            )
        ]
        projects = [row["project_name"] for row in conn.execute("SELECT DISTINCT project_name FROM work_entries ORDER BY project_name COLLATE NOCASE")]
    return {"months": months, "projects": projects}


def fetch_entries(sqlite_db_path: Path, month: str, project: str) -> list[dict]:
    if USE_POSTGRES:
        query = """
            SELECT id, work_date, employee_name, project_name, task_name, hours_spent
            FROM work_entries
            WHERE substring(work_date, 1, 7) = %s
        """
        params: list[Any] = [month]
        if project and project != ALL_PROJECTS:
            query += " AND project_name = %s"
            params.append(project)
        query += " ORDER BY work_date DESC, id DESC"
        with psycopg.connect(SUPABASE_DB_URL) as conn:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(query, params)
                rows = cur.fetchall()
        return [
            {
                "id": int(row["id"]),
                "work_date": row["work_date"],
                "employee_name": row["employee_name"],
                "project_name": row["project_name"],
                "task_name": row["task_name"],
                "hours_spent": float(row["hours_spent"]),
            }
            for row in rows
        ]

    query = """
        SELECT id, work_date, employee_name, project_name, task_name, hours_spent
        FROM work_entries
        WHERE substr(work_date, 1, 7) = ?
    """
    params2: list[Any] = [month]
    if project and project != ALL_PROJECTS:
        query += " AND project_name = ?"
        params2.append(project)
    query += " ORDER BY work_date DESC, id DESC"
    with _sqlite_connection(sqlite_db_path) as conn:
        rows = conn.execute(query, params2).fetchall()
    return [
        {
            "id": int(row["id"]),
            "work_date": row["work_date"],
            "employee_name": row["employee_name"],
            "project_name": row["project_name"],
            "task_name": row["task_name"],
            "hours_spent": float(row["hours_spent"]),
        }
        for row in rows
    ]


def fetch_all_entries(sqlite_db_path: Path) -> list[dict]:
    if USE_POSTGRES:
        with psycopg.connect(SUPABASE_DB_URL) as conn:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(
                    """
                    SELECT id, work_date, employee_name, project_name, task_name, hours_spent
                    FROM work_entries
                    ORDER BY work_date DESC, id DESC
                    """
                )
                rows = cur.fetchall()
        return [
            {
                "id": int(row["id"]),
                "work_date": row["work_date"],
                "employee_name": row["employee_name"],
                "project_name": row["project_name"],
                "task_name": row["task_name"],
                "hours_spent": float(row["hours_spent"]),
            }
            for row in rows
        ]

    with _sqlite_connection(sqlite_db_path) as conn:
        rows = conn.execute(
            """
            SELECT id, work_date, employee_name, project_name, task_name, hours_spent
            FROM work_entries
            ORDER BY work_date DESC, id DESC
            """
        ).fetchall()
    return [
        {
            "id": int(row["id"]),
            "work_date": row["work_date"],
            "employee_name": row["employee_name"],
            "project_name": row["project_name"],
            "task_name": row["task_name"],
            "hours_spent": float(row["hours_spent"]),
        }
        for row in rows
    ]


def fetch_monthly_summary(sqlite_db_path: Path) -> list[list]:
    if USE_POSTGRES:
        with psycopg.connect(SUPABASE_DB_URL) as conn:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(
                    """
                    SELECT substring(work_date, 1, 7) AS month_key, COUNT(*) AS entry_count, ROUND(SUM(hours_spent)::numeric, 2) AS total_hours
                    FROM work_entries
                    GROUP BY substring(work_date, 1, 7)
                    ORDER BY month_key DESC
                    """
                )
                rows = cur.fetchall()
        return [[row["month_key"], int(row["entry_count"]), float(row["total_hours"] or 0)] for row in rows]

    with _sqlite_connection(sqlite_db_path) as conn:
        rows = conn.execute(
            """
            SELECT substr(work_date, 1, 7) AS month_key, COUNT(*) AS entry_count, ROUND(SUM(hours_spent), 2) AS total_hours
            FROM work_entries
            GROUP BY substr(work_date, 1, 7)
            ORDER BY month_key DESC
            """
        ).fetchall()
    return [[row["month_key"], int(row["entry_count"]), float(row["total_hours"] or 0)] for row in rows]


def fetch_history_rows(sqlite_db_path: Path, month: str | None, project: str | None, limit: int = 40) -> list[dict]:
    if USE_POSTGRES:
        query = """
            SELECT id, entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, changed_at
            FROM change_history
            WHERE 1 = 1
        """
        params: list[Any] = []
        if month:
            query += " AND substring(work_date, 1, 7) = %s"
            params.append(month)
        if project and project != ALL_PROJECTS:
            query += " AND project_name = %s"
            params.append(project)
        query += " ORDER BY changed_at DESC, id DESC LIMIT %s"
        params.append(limit)
        with psycopg.connect(SUPABASE_DB_URL) as conn:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(query, params)
                rows = cur.fetchall()
        return [
            {
                "id": int(row["id"]),
                "entry_id": int(row["entry_id"]) if row["entry_id"] is not None else None,
                "action": row["action"],
                "changed_by": row["changed_by"],
                "employee_name": row["employee_name"],
                "work_date": row["work_date"],
                "project_name": row["project_name"],
                "task_name": row["task_name"],
                "hours_spent": float(row["hours_spent"] or 0),
                "changed_at": str(row["changed_at"]),
            }
            for row in rows
        ]

    query2 = """
        SELECT id, entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, changed_at
        FROM change_history
        WHERE 1 = 1
    """
    params2: list[Any] = []
    if month:
        query2 += " AND substr(work_date, 1, 7) = ?"
        params2.append(month)
    if project and project != ALL_PROJECTS:
        query2 += " AND project_name = ?"
        params2.append(project)
    query2 += " ORDER BY changed_at DESC, id DESC LIMIT ?"
    params2.append(limit)
    with _sqlite_connection(sqlite_db_path) as conn:
        rows = conn.execute(query2, params2).fetchall()
    return [dict(row) for row in rows]


def create_entry(sqlite_db_path: Path, values: dict, operator: str) -> int:
    if USE_POSTGRES:
        with psycopg.connect(SUPABASE_DB_URL) as conn:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(
                    """
                    INSERT INTO work_entries (employee_name, work_date, project_name, task_name, hours_spent)
                    VALUES (%s, %s, %s, %s, %s)
                    RETURNING id
                    """,
                    (values["employee_name"], values["work_date"], values["project_name"], values["task_name"], values["hours_spent"]),
                )
                entry_id = int(cur.fetchone()["id"])
                _insert_change_pg(cur, entry_id, "CREATE", operator, None, values)
            conn.commit()
        return entry_id

    with _sqlite_connection(sqlite_db_path) as conn:
        conn.execute(
            """
            INSERT INTO work_entries (employee_name, work_date, project_name, task_name, hours_spent)
            VALUES (?, ?, ?, ?, ?)
            """,
            (values["employee_name"], values["work_date"], values["project_name"], values["task_name"], values["hours_spent"]),
        )
        entry_id = int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])
        _insert_change_sqlite(conn, entry_id, "CREATE", operator, None, values)
        conn.commit()
    return entry_id


def update_entry(sqlite_db_path: Path, entry_id: int, values: dict, operator: str) -> bool:
    if USE_POSTGRES:
        with psycopg.connect(SUPABASE_DB_URL) as conn:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(
                    """
                    SELECT employee_name, work_date, project_name, task_name, hours_spent
                    FROM work_entries
                    WHERE id = %s
                    """,
                    (entry_id,),
                )
                old_row = cur.fetchone()
                if old_row is None:
                    return False
                old_values = {
                    "employee_name": old_row["employee_name"],
                    "work_date": old_row["work_date"],
                    "project_name": old_row["project_name"],
                    "task_name": old_row["task_name"],
                    "hours_spent": float(old_row["hours_spent"]),
                }
                cur.execute(
                    """
                    UPDATE work_entries
                    SET employee_name = %s, work_date = %s, project_name = %s, task_name = %s, hours_spent = %s
                    WHERE id = %s
                    """,
                    (values["employee_name"], values["work_date"], values["project_name"], values["task_name"], values["hours_spent"], entry_id),
                )
                _insert_change_pg(cur, entry_id, "UPDATE", operator, old_values, values)
            conn.commit()
        return True

    with _sqlite_connection(sqlite_db_path) as conn:
        row = conn.execute(
            """
            SELECT employee_name, work_date, project_name, task_name, hours_spent
            FROM work_entries
            WHERE id = ?
            """,
            (entry_id,),
        ).fetchone()
        if row is None:
            return False
        old_values = {
            "employee_name": row["employee_name"],
            "work_date": row["work_date"],
            "project_name": row["project_name"],
            "task_name": row["task_name"],
            "hours_spent": float(row["hours_spent"]),
        }
        conn.execute(
            """
            UPDATE work_entries
            SET employee_name = ?, work_date = ?, project_name = ?, task_name = ?, hours_spent = ?
            WHERE id = ?
            """,
            (values["employee_name"], values["work_date"], values["project_name"], values["task_name"], values["hours_spent"], entry_id),
        )
        _insert_change_sqlite(conn, entry_id, "UPDATE", operator, old_values, values)
        conn.commit()
    return True


def delete_entry(sqlite_db_path: Path, entry_id: int, operator: str) -> bool:
    if USE_POSTGRES:
        with psycopg.connect(SUPABASE_DB_URL) as conn:
            with conn.cursor(row_factory=dict_row) as cur:
                cur.execute(
                    """
                    SELECT employee_name, work_date, project_name, task_name, hours_spent
                    FROM work_entries
                    WHERE id = %s
                    """,
                    (entry_id,),
                )
                old_row = cur.fetchone()
                if old_row is None:
                    return False
                old_values = {
                    "employee_name": old_row["employee_name"],
                    "work_date": old_row["work_date"],
                    "project_name": old_row["project_name"],
                    "task_name": old_row["task_name"],
                    "hours_spent": float(old_row["hours_spent"]),
                }
                cur.execute("DELETE FROM work_entries WHERE id = %s", (entry_id,))
                _insert_change_pg(cur, entry_id, "DELETE", operator, old_values, None)
            conn.commit()
        return True

    with _sqlite_connection(sqlite_db_path) as conn:
        row = conn.execute(
            """
            SELECT employee_name, work_date, project_name, task_name, hours_spent
            FROM work_entries
            WHERE id = ?
            """,
            (entry_id,),
        ).fetchone()
        if row is None:
            return False
        old_values = {
            "employee_name": row["employee_name"],
            "work_date": row["work_date"],
            "project_name": row["project_name"],
            "task_name": row["task_name"],
            "hours_spent": float(row["hours_spent"]),
        }
        conn.execute("DELETE FROM work_entries WHERE id = ?", (entry_id,))
        _insert_change_sqlite(conn, entry_id, "DELETE", operator, old_values, None)
        conn.commit()
    return True


def _insert_change_sqlite(
    conn: sqlite3.Connection,
    entry_id: int | None,
    action: str,
    operator: str,
    old_values: dict | None,
    new_values: dict | None,
) -> None:
    current = new_values or old_values or {}
    conn.execute(
        """
        INSERT INTO change_history (
            entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, old_values, new_values
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            entry_id,
            action,
            operator,
            current.get("employee_name"),
            current.get("work_date"),
            current.get("project_name"),
            current.get("task_name"),
            current.get("hours_spent"),
            json.dumps(old_values or {}, ensure_ascii=False),
            json.dumps(new_values or {}, ensure_ascii=False),
        ),
    )


def _insert_change_pg(
    cur,
    entry_id: int | None,
    action: str,
    operator: str,
    old_values: dict | None,
    new_values: dict | None,
) -> None:
    current = new_values or old_values or {}
    cur.execute(
        """
        INSERT INTO change_history (
            entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, old_values, new_values
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s::jsonb, %s::jsonb)
        """,
        (
            entry_id,
            action,
            operator,
            current.get("employee_name"),
            current.get("work_date"),
            current.get("project_name"),
            current.get("task_name"),
            current.get("hours_spent"),
            json.dumps(old_values or {}, ensure_ascii=False),
            json.dumps(new_values or {}, ensure_ascii=False),
        ),
    )
