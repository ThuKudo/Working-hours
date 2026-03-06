# Working Hours Web - Deploy Notes

## 1) Stable mode on Vercel (Supabase Postgres)
This app now supports two backends:
- Local development fallback: SQLite (`worklog.db`)
- Production stable: Postgres via `SUPABASE_DB_URL` (or `DATABASE_URL`)

If `SUPABASE_DB_URL` is set to a Postgres URL, the app will use Supabase and data is persistent.

## 2) Required environment variables on Vercel
In Vercel project settings, add:
- `SUPABASE_DB_URL` = your Supabase Postgres connection string

Then redeploy. The app will auto-create tables (`work_entries`, `change_history`) on first run.

## 3) Optional one-time migration from SQLite to Supabase
Run locally in this folder:

```powershell
cd "d:\03. LEARNING\0. Python\00. CODE TEST\Kudo\working-hours-web-vercel"
python -m pip install -r requirements.txt
$env:SUPABASE_DB_URL="<your_supabase_postgres_url>"
python migrate_sqlite_to_supabase.py
```

## 4) Push to GitHub
Run in this folder:

```powershell
cd "d:\03. LEARNING\0. Python\00. CODE TEST\Kudo\working-hours-web-vercel"
git add .
git commit -m "Migrate Working Hours backend to Supabase-ready mode"
git push origin main
```

## 5) Deploy on Vercel
1. Login Vercel: https://vercel.com
2. Open your project connected to this GitHub repo
3. Add `SUPABASE_DB_URL` in Environment Variables
4. Trigger a new deploy from latest `main`
