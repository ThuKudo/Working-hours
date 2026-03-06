from __future__ import annotations

import os
import sqlite3
from pathlib import Path

import psycopg
from psycopg.rows import dict_row


BASE_DIR = Path(__file__).resolve().parent
SQLITE_PATH = BASE_DIR / "worklog.db"
SUPABASE_DB_URL = os.environ.get("SUPABASE_DB_URL") or os.environ.get("DATABASE_URL")


def ensure_pg_schema(conn) -> None:
    with conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS work_entries (
                id BIGSERIAL PRIMARY KEY,
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
    conn.commit()


def main() -> None:
    if not SUPABASE_DB_URL:
        raise SystemExit("Missing SUPABASE_DB_URL (or DATABASE_URL).")
    if not SQLITE_PATH.exists():
        raise SystemExit(f"SQLite file not found: {SQLITE_PATH}")

    sqlite_conn = sqlite3.connect(SQLITE_PATH)
    sqlite_conn.row_factory = sqlite3.Row

    with sqlite_conn:
        entries = sqlite_conn.execute(
            """
            SELECT id, employee_name, work_date, project_name, task_name, hours_spent
            FROM work_entries
            ORDER BY id ASC
            """
        ).fetchall()
        history_rows = sqlite_conn.execute(
            """
            SELECT entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, old_values, new_values
            FROM change_history
            ORDER BY id ASC
            """
        ).fetchall()

    with psycopg.connect(SUPABASE_DB_URL) as pg_conn:
        ensure_pg_schema(pg_conn)
        with pg_conn.cursor(row_factory=dict_row) as cur:
            cur.execute("SELECT COUNT(*) AS c FROM work_entries")
            existing_entries = int(cur.fetchone()["c"])
            cur.execute("SELECT COUNT(*) AS c FROM change_history")
            existing_history = int(cur.fetchone()["c"])
            if existing_entries > 0 or existing_history > 0:
                raise SystemExit("Target Postgres is not empty. Stop to avoid duplicate migration.")

            id_map: dict[int, int] = {}
            for row in entries:
                cur.execute(
                    """
                    INSERT INTO work_entries (employee_name, work_date, project_name, task_name, hours_spent)
                    VALUES (%s, %s, %s, %s, %s)
                    RETURNING id
                    """,
                    (
                        row["employee_name"],
                        row["work_date"],
                        row["project_name"],
                        row["task_name"],
                        float(row["hours_spent"]),
                    ),
                )
                new_id = int(cur.fetchone()["id"])
                id_map[int(row["id"])] = new_id

            for row in history_rows:
                old_entry_id = row["entry_id"]
                mapped_entry_id = id_map.get(int(old_entry_id)) if old_entry_id is not None else None
                cur.execute(
                    """
                    INSERT INTO change_history (
                        entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, old_values, new_values
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s::jsonb, %s::jsonb)
                    """,
                    (
                        mapped_entry_id,
                        row["action"],
                        row["changed_by"],
                        row["employee_name"],
                        row["work_date"],
                        row["project_name"],
                        row["task_name"],
                        float(row["hours_spent"] or 0),
                        row["old_values"] or "{}",
                        row["new_values"] or "{}",
                    ),
                )
        pg_conn.commit()

    print(f"Migrated {len(entries)} entries and {len(history_rows)} history rows.")


if __name__ == "__main__":
    main()
