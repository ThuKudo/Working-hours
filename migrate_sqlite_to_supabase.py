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
            SELECT id, entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, old_values, new_values
            FROM change_history
            ORDER BY id ASC
            """
        ).fetchall()

    with psycopg.connect(SUPABASE_DB_URL) as pg_conn:
        ensure_pg_schema(pg_conn)
        with pg_conn.cursor(row_factory=dict_row) as cur:
            id_map: dict[int, int] = {}
            inserted_entries = 0
            updated_entries = 0
            for row in entries:
                cur.execute(
                    """
                    INSERT INTO work_entries (legacy_sqlite_id, employee_name, work_date, project_name, task_name, hours_spent)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (legacy_sqlite_id) DO UPDATE
                    SET employee_name = EXCLUDED.employee_name,
                        work_date = EXCLUDED.work_date,
                        project_name = EXCLUDED.project_name,
                        task_name = EXCLUDED.task_name,
                        hours_spent = EXCLUDED.hours_spent
                    RETURNING id, (xmax = 0) AS inserted
                    """,
                    (
                        int(row["id"]),
                        row["employee_name"],
                        row["work_date"],
                        row["project_name"],
                        row["task_name"],
                        float(row["hours_spent"]),
                    ),
                )
                result = cur.fetchone()
                new_id = int(result["id"])
                if bool(result["inserted"]):
                    inserted_entries += 1
                else:
                    updated_entries += 1
                id_map[int(row["id"])] = new_id

            inserted_history = 0
            updated_history = 0
            for row in history_rows:
                old_entry_id = row["entry_id"]
                mapped_entry_id = id_map.get(int(old_entry_id)) if old_entry_id is not None else None
                cur.execute(
                    """
                    INSERT INTO change_history (
                        legacy_sqlite_history_id, entry_id, action, changed_by, employee_name, work_date, project_name, task_name, hours_spent, old_values, new_values
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s::jsonb, %s::jsonb)
                    ON CONFLICT (legacy_sqlite_history_id) DO UPDATE
                    SET entry_id = EXCLUDED.entry_id,
                        action = EXCLUDED.action,
                        changed_by = EXCLUDED.changed_by,
                        employee_name = EXCLUDED.employee_name,
                        work_date = EXCLUDED.work_date,
                        project_name = EXCLUDED.project_name,
                        task_name = EXCLUDED.task_name,
                        hours_spent = EXCLUDED.hours_spent,
                        old_values = EXCLUDED.old_values,
                        new_values = EXCLUDED.new_values
                    RETURNING (xmax = 0) AS inserted
                    """,
                    (
                        int(row["id"]),
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
                result = cur.fetchone()
                if bool(result["inserted"]):
                    inserted_history += 1
                else:
                    updated_history += 1
        pg_conn.commit()

    print(
        "Migration synced successfully."
        f" Entries inserted={inserted_entries}, updated={updated_entries};"
        f" History inserted={inserted_history}, updated={updated_history}."
    )


if __name__ == "__main__":
    main()
