# Working Hours Web - Deploy Notes

## 1) Files prepared in this folder
- `webapp.py`
- `audit_utils.py`
- `export_utils.py`
- `worklog.db` (optional seed data)

## 2) Important limitation on Vercel
This app uses local SQLite (`worklog.db`) for create/update/delete operations.
On Vercel serverless, local filesystem is ephemeral and not persistent for writes.

That means:
- Reads from bundled file may work.
- Writes (save/edit/delete) are not reliable/persistent.

For real use, move data to a hosted DB (Supabase Postgres, Neon, PlanetScale, etc).

## 3) Push to GitHub
Run in this folder:

```powershell
cd "d:\03. LEARNING\0. Python\00. CODE TEST\Kudo\working-hours-web-vercel"
git init
git add .
git commit -m "Init working-hours web deploy folder"
git branch -M main
git remote add origin <YOUR_GITHUB_REPO_URL>
git push -u origin main
```

## 4) Deploy on Vercel
1. Login Vercel: https://vercel.com
2. Click `Add New...` -> `Project`
3. Import your GitHub repository
4. Framework Preset: `Other`
5. Root Directory: this repo root
6. Click `Deploy`

## 5) Recommended production path
- Keep Vercel for frontend.
- Replace SQLite with hosted DB.
- Refactor `webapp.py` from local `http.server` style to serverless-compatible API routes.
