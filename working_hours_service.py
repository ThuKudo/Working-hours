import json
import os
import sqlite3
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from flask import Flask, Response, jsonify, request

from audit_utils import default_operator, ensure_schema, fetch_history, log_change, snapshot_from_row
from export_utils import export_excel_report


ALL_PROJECTS = "All Projects"
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = (Path("/tmp") / "worklog.db") if os.environ.get("VERCEL") else (BASE_DIR / "worklog.db")


app = Flask(__name__)


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    ensure_schema(conn)
    return conn


def fetch_filters(conn: sqlite3.Connection) -> dict:
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
    projects = [
        row["project_name"]
        for row in conn.execute(
            "SELECT DISTINCT project_name FROM work_entries ORDER BY project_name COLLATE NOCASE"
        )
    ]
    return {"months": months, "projects": projects}


def fetch_entries(conn: sqlite3.Connection, month: str, project: str) -> list[dict]:
    query = """
        SELECT id, work_date, employee_name, project_name, task_name, hours_spent
        FROM work_entries
        WHERE substr(work_date, 1, 7) = ?
    """
    params: list[str] = [month]
    if project and project != ALL_PROJECTS:
        query += " AND project_name = ?"
        params.append(project)
    query += " ORDER BY work_date DESC, id DESC"
    rows = conn.execute(query, params).fetchall()
    return [
        {
            "id": row["id"],
            "work_date": row["work_date"],
            "employee_name": row["employee_name"],
            "project_name": row["project_name"],
            "task_name": row["task_name"],
            "hours_spent": float(row["hours_spent"]),
        }
        for row in rows
    ]


def build_summary(entries: list[dict]) -> dict:
    by_project: dict[str, float] = defaultdict(float)
    by_employee: dict[str, float] = defaultdict(float)
    total_hours = 0.0
    for entry in entries:
        hours = float(entry["hours_spent"])
        total_hours += hours
        by_project[entry["project_name"]] += hours
        by_employee[entry["employee_name"]] += hours
    top_project = ""
    if by_project:
        name, hours = max(by_project.items(), key=lambda item: item[1])
        top_project = f"{name} ({hours:.1f}h)"
    return {
        "total_hours": total_hours,
        "entry_count": len(entries),
        "top_project": top_project,
        "by_project": dict(sorted(by_project.items(), key=lambda item: item[1], reverse=True)),
        "by_employee": dict(sorted(by_employee.items(), key=lambda item: item[0].lower())),
    }


def validate_payload(payload: dict) -> tuple[bool, str]:
    required = ["employee_name", "work_date", "project_name", "task_name", "hours_spent", "operator_name"]
    if not all(str(payload.get(field, "")).strip() for field in required):
        return False, "Please complete all fields."
    try:
        datetime.strptime(str(payload["work_date"]), "%Y-%m-%d")
    except ValueError:
        return False, "Work Date must use the YYYY-MM-DD format."
    try:
        hours = float(payload["hours_spent"])
        if hours <= 0:
            raise ValueError
    except (TypeError, ValueError):
        return False, "Hours must be a positive number."
    return True, ""


@app.get("/")
def index() -> Response:
    html = HTML_PAGE.replace("__DEFAULT_OPERATOR__", json.dumps(default_operator())[1:-1])
    return Response(html, mimetype="text/html; charset=utf-8")


@app.get("/api/filters")
def api_filters():
    with get_connection() as conn:
        return jsonify(fetch_filters(conn))


@app.get("/api/dashboard")
def api_dashboard():
    month = request.args.get("month") or date.today().strftime("%Y-%m")
    project = request.args.get("project") or ALL_PROJECTS
    with get_connection() as conn:
        entries = fetch_entries(conn, month, project)
        history = [dict(row) for row in fetch_history(conn, month=month, project=project, limit=40)]
    return jsonify({"entries": entries, "summary": build_summary(entries), "history": history})


@app.post("/api/entries")
def api_create_entry():
    payload = request.get_json(silent=True) or {}
    ok, error = validate_payload(payload)
    if not ok:
        return jsonify({"error": error}), 400

    operator = str(payload["operator_name"]).strip()
    values = {
        "employee_name": str(payload["employee_name"]).strip(),
        "work_date": str(payload["work_date"]).strip(),
        "project_name": str(payload["project_name"]).strip(),
        "task_name": str(payload["task_name"]).strip(),
        "hours_spent": round(float(payload["hours_spent"]), 2),
    }

    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO work_entries (employee_name, work_date, project_name, task_name, hours_spent)
            VALUES (?, ?, ?, ?, ?)
            """,
            (values["employee_name"], values["work_date"], values["project_name"], values["task_name"], values["hours_spent"]),
        )
        entry_id = int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])
        log_change(conn, entry_id=entry_id, action="CREATE", changed_by=operator, new_values=values)
        conn.commit()
    return jsonify({"status": "ok"}), 201


@app.put("/api/entries/<int:entry_id>")
def api_update_entry(entry_id: int):
    payload = request.get_json(silent=True) or {}
    ok, error = validate_payload(payload)
    if not ok:
        return jsonify({"error": error}), 400

    operator = str(payload["operator_name"]).strip()
    values = {
        "employee_name": str(payload["employee_name"]).strip(),
        "work_date": str(payload["work_date"]).strip(),
        "project_name": str(payload["project_name"]).strip(),
        "task_name": str(payload["task_name"]).strip(),
        "hours_spent": round(float(payload["hours_spent"]), 2),
    }

    with get_connection() as conn:
        old_row = conn.execute(
            """
            SELECT employee_name, work_date, project_name, task_name, hours_spent
            FROM work_entries
            WHERE id = ?
            """,
            (entry_id,),
        ).fetchone()
        if old_row is None:
            return jsonify({"error": "Entry not found."}), 404
        conn.execute(
            """
            UPDATE work_entries
            SET employee_name = ?, work_date = ?, project_name = ?, task_name = ?, hours_spent = ?
            WHERE id = ?
            """,
            (values["employee_name"], values["work_date"], values["project_name"], values["task_name"], values["hours_spent"], entry_id),
        )
        log_change(
            conn,
            entry_id=entry_id,
            action="UPDATE",
            changed_by=operator,
            old_values=snapshot_from_row(old_row),
            new_values=values,
        )
        conn.commit()
    return jsonify({"status": "ok"})


@app.delete("/api/entries/<int:entry_id>")
def api_delete_entry(entry_id: int):
    payload = request.get_json(silent=True) or {}
    operator = str(payload.get("operator_name", "")).strip()
    if not operator:
        return jsonify({"error": "Please enter the operator name."}), 400

    with get_connection() as conn:
        old_row = conn.execute(
            """
            SELECT employee_name, work_date, project_name, task_name, hours_spent
            FROM work_entries
            WHERE id = ?
            """,
            (entry_id,),
        ).fetchone()
        if old_row is None:
            return jsonify({"error": "Entry not found."}), 404
        conn.execute("DELETE FROM work_entries WHERE id = ?", (entry_id,))
        log_change(
            conn,
            entry_id=entry_id,
            action="DELETE",
            changed_by=operator,
            old_values=snapshot_from_row(old_row),
        )
        conn.commit()
    return jsonify({"status": "ok"})


@app.get("/api/export.xlsx")
def api_export_excel():
    month = request.args.get("month") or date.today().strftime("%Y-%m")
    project = request.args.get("project") or ALL_PROJECTS
    file_bytes = export_excel_report(DB_PATH, month=month, project=project)
    file_name = f"working-hours-report-{date.today().isoformat()}.xlsx"
    return Response(
        file_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
    )


HTML_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Working Hours</title>
  <style>
    body { font-family: Segoe UI, sans-serif; margin: 0; background: #f5f1e8; color: #20364b; }
    .wrap { max-width: 1280px; margin: 0 auto; padding: 20px; }
    .grid { display: grid; grid-template-columns: 340px 1fr; gap: 16px; }
    .card { background: #fffdf8; border: 1px solid #ded6ca; border-radius: 12px; padding: 16px; }
    .field { margin-bottom: 10px; }
    label { display:block; font-size: 13px; margin-bottom: 4px; color: #6c6258; }
    input, textarea, select, button { width: 100%; padding: 10px; border-radius: 8px; border: 1px solid #cfcbc2; font: inherit; }
    button { cursor: pointer; background: #2d5b8a; color: #fff; border: 0; }
    .row { display: grid; grid-template-columns: 1fr 1fr 130px; gap: 10px; align-items: end; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border-bottom: 1px solid #ece5db; padding: 10px; text-align: left; }
    .charts { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .bar-track { background:#efe8dd; border-radius:999px; height: 12px; overflow: hidden; }
    .bar-fill { height: 100%; }
    .bar-row { display:grid; grid-template-columns:100px 1fr 50px; gap:8px; margin: 6px 0; }
    .pie { width: 220px; height: 220px; border-radius: 50%; border: 10px solid #f1eadf; }
    .legend-item { margin: 6px 0; font-size: 13px; }
    .status { min-height: 18px; margin-top: 8px; font-size: 13px; color: #6c6258; }
    .status.error { color: #b65050; }
    @media (max-width: 1020px) { .grid { grid-template-columns: 1fr; } .charts { grid-template-columns: 1fr; } .row { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
<div class="wrap">
  <h1>Working Hours</h1>
  <div class="grid">
    <div class="card">
      <h3 id="formTitle">Add Work Entry</h3>
      <form id="entryForm">
        <div class="field"><label>Employee</label><input id="employee_name" name="employee_name" required /></div>
        <div class="field"><label>Work Date</label><input id="work_date" name="work_date" type="date" required /></div>
        <div class="field"><label>Project</label><input id="project_name" name="project_name" required /></div>
        <div class="field"><label>Task</label><textarea id="task_name" name="task_name" required></textarea></div>
        <div class="field"><label>Hours</label><input id="hours_spent" name="hours_spent" type="number" min="0.25" step="0.25" required /></div>
        <div class="field"><label>Operator</label><input id="operator_name" name="operator_name" required /></div>
        <button id="submitButton" type="submit">Save Entry</button>
        <button id="cancelEditButton" type="button" style="display:none; margin-top:8px; background:#8b8175;">Cancel Edit</button>
        <div id="formStatus" class="status"></div>
      </form>
    </div>
    <div>
      <div class="card" style="margin-bottom: 10px;">
        <div class="row">
          <div><label>Month</label><select id="monthFilter"></select></div>
          <div><label>Project</label><select id="projectFilter"></select></div>
          <div><button id="refreshButton" type="button">Refresh</button></div>
        </div>
        <button id="exportButton" type="button" style="margin-top:8px;">Export Excel</button>
      </div>
      <div class="card" style="margin-bottom: 10px;">
        <div id="summary"></div>
      </div>
      <div class="card" style="margin-bottom: 10px;">
        <h3>Entries</h3>
        <table>
          <thead><tr><th>Date</th><th>Employee</th><th>Project</th><th>Task</th><th>Hours</th><th>Action</th></tr></thead>
          <tbody id="entriesBody"></tbody>
        </table>
      </div>
      <div class="charts">
        <div class="card"><h3>Hours by Project</h3><div id="projectBars"></div></div>
        <div class="card"><h3>Hours Share by Employee</h3><div id="employeePie" class="pie"></div><div id="employeeLegend"></div></div>
      </div>
      <div class="card" style="margin-top:10px;">
        <h3>Recent Change History</h3>
        <table>
          <thead><tr><th>Changed At</th><th>Action</th><th>Changed By</th><th>Employee</th><th>Project</th><th>Work Date</th><th>Hours</th></tr></thead>
          <tbody id="historyBody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>
<script>
const state = { month:"", project:"All Projects", editingId:null, colors:["#2d5b8a","#c97b63","#7b9e87","#d5a34d","#8a6f9e","#5d8c8c"] };
const entriesBody = document.getElementById("entriesBody");
const historyBody = document.getElementById("historyBody");
const formStatus = document.getElementById("formStatus");
const formTitle = document.getElementById("formTitle");
const submitButton = document.getElementById("submitButton");
const cancelEditButton = document.getElementById("cancelEditButton");
document.getElementById("work_date").value = new Date().toISOString().slice(0,10);
document.getElementById("operator_name").value = "__DEFAULT_OPERATOR__";

async function fetchJson(url){ const r=await fetch(url); return r.json(); }
function resetForm(){ state.editingId=null; document.getElementById("entryForm").reset(); document.getElementById("work_date").value=new Date().toISOString().slice(0,10); document.getElementById("operator_name").value="__DEFAULT_OPERATOR__"; formTitle.textContent="Add Work Entry"; submitButton.textContent="Save Entry"; cancelEditButton.style.display="none"; formStatus.textContent=""; formStatus.className="status"; }
cancelEditButton.addEventListener("click", resetForm);

document.getElementById("entryForm").addEventListener("submit", async (e)=>{
  e.preventDefault();
  const fd = new FormData(e.target); const payload = Object.fromEntries(fd.entries()); payload.hours_spent = Number(payload.hours_spent);
  if(!payload.operator_name?.trim()){ formStatus.textContent="Please enter the operator name."; formStatus.className="status error"; return; }
  const isEdit = state.editingId !== null;
  if(isEdit && !confirm("Do you want to update the selected entry?")) return;
  const resp = await fetch(isEdit?`/api/entries/${state.editingId}`:"/api/entries",{method:isEdit?"PUT":"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
  const res = await resp.json();
  if(!resp.ok){ formStatus.textContent=res.error||"Could not save."; formStatus.className="status error"; return; }
  formStatus.textContent=isEdit?"Entry updated.":"Entry saved."; formStatus.className="status"; resetForm(); state.month=payload.work_date.slice(0,7); state.project="All Projects"; await initFilters(); await loadDashboard();
});

document.getElementById("refreshButton").addEventListener("click", loadDashboard);
document.getElementById("exportButton").addEventListener("click", ()=>{ const q=new URLSearchParams({month:state.month,project:state.project}); window.location.href=`/api/export.xlsx?${q.toString()}`; });
document.getElementById("monthFilter").addEventListener("change", async e=>{ state.month=e.target.value; await loadDashboard(); });
document.getElementById("projectFilter").addEventListener("change", async e=>{ state.project=e.target.value; await loadDashboard(); });

async function initFilters(){
  const f=await fetchJson("/api/filters");
  const months=f.months.length?f.months:[new Date().toISOString().slice(0,7)];
  const projects=f.projects.length?["All Projects",...f.projects]:["All Projects"];
  if(!state.month||!months.includes(state.month)) state.month=months[0];
  if(!projects.includes(state.project)) state.project="All Projects";
  document.getElementById("monthFilter").innerHTML = months.map(x=>`<option value="${x}">${x}</option>`).join("");
  document.getElementById("projectFilter").innerHTML = projects.map(x=>`<option value="${x}">${x}</option>`).join("");
  document.getElementById("monthFilter").value = state.month;
  document.getElementById("projectFilter").value = state.project;
}

async function loadDashboard(){
  const q=new URLSearchParams({month:state.month,project:state.project});
  const data=await fetchJson(`/api/dashboard?${q.toString()}`);
  renderSummary(data.summary); renderTable(data.entries); renderBars(data.summary.by_project); renderPie(data.summary.by_employee); renderHistory(data.history||[]);
}

function renderSummary(s){ document.getElementById("summary").innerHTML=`<strong>Monthly Hours:</strong> ${s.total_hours.toFixed(1)} | <strong>Tasks:</strong> ${s.entry_count} | <strong>Top Project:</strong> ${s.top_project||"No data"}`; }

function renderTable(entries){
  if(!entries.length){ entriesBody.innerHTML='<tr><td colspan="6">No data for current filter.</td></tr>'; return; }
  entriesBody.innerHTML = entries.map(e=>`<tr><td>${e.work_date}</td><td>${e.employee_name}</td><td>${e.project_name}</td><td>${e.task_name}</td><td>${e.hours_spent.toFixed(2)}</td><td><button data-edit="${e.id}" style="padding:6px;">Edit</button> <button data-del="${e.id}" style="padding:6px; background:#b65050;">Delete</button></td></tr>`).join("");
  entriesBody.querySelectorAll("[data-edit]").forEach(btn=>btn.onclick=()=>startEdit(Number(btn.dataset.edit),entries));
  entriesBody.querySelectorAll("[data-del]").forEach(btn=>btn.onclick=()=>deleteEntry(Number(btn.dataset.del)));
}

function startEdit(id, entries){
  const e=entries.find(x=>x.id===id); if(!e) return;
  state.editingId=e.id; document.getElementById("employee_name").value=e.employee_name; document.getElementById("work_date").value=e.work_date; document.getElementById("project_name").value=e.project_name; document.getElementById("task_name").value=e.task_name; document.getElementById("hours_spent").value=e.hours_spent;
  formTitle.textContent="Edit Work Entry"; submitButton.textContent="Update Entry"; cancelEditButton.style.display="block"; formStatus.textContent="Editing selected entry."; formStatus.className="status";
}

async function deleteEntry(id){
  const op=document.getElementById("operator_name").value.trim();
  if(!op){ formStatus.textContent="Please enter the operator name."; formStatus.className="status error"; return; }
  if(!confirm("Do you want to delete the selected entry?")) return;
  const resp=await fetch(`/api/entries/${id}`,{method:"DELETE",headers:{"Content-Type":"application/json"},body:JSON.stringify({operator_name:op})});
  const res=await resp.json();
  if(!resp.ok){ formStatus.textContent=res.error||"Delete failed."; formStatus.className="status error"; return; }
  if(state.editingId===id) resetForm();
  formStatus.textContent="Entry deleted."; formStatus.className="status";
  await loadDashboard();
}

function renderBars(byProject){
  const box=document.getElementById("projectBars"); const names=Object.keys(byProject||{});
  if(!names.length){ box.innerHTML="<div>No data.</div>"; return; }
  const max=Math.max(...Object.values(byProject),1);
  box.innerHTML = names.map((n,i)=>{ const v=byProject[n]; const w=(v/max)*100; const c=state.colors[i%state.colors.length]; return `<div class="bar-row"><div>${n}</div><div class="bar-track"><div class="bar-fill" style="width:${w}%; background:${c};"></div></div><div>${v.toFixed(1)}</div></div>`; }).join("");
}

function renderPie(byEmployee){
  const pie=document.getElementById("employeePie"); const legend=document.getElementById("employeeLegend"); const names=Object.keys(byEmployee||{});
  if(!names.length){ pie.style.background="#efe8dd"; legend.innerHTML="<div>No data.</div>"; return; }
  const total=Object.values(byEmployee).reduce((a,b)=>a+b,0)||1; let start=0;
  const seg=names.map((n,i)=>{ const v=byEmployee[n]; const pct=(v/total)*100; const end=start+pct; const c=state.colors[i%state.colors.length]; const s=`${c} ${start.toFixed(2)}% ${end.toFixed(2)}%`; start=end; return s; });
  pie.style.background=`conic-gradient(${seg.join(",")})`;
  legend.innerHTML = names.map((n,i)=>{ const v=byEmployee[n]; const pct=(v/total)*100; return `<div class="legend-item" style="color:${state.colors[i%state.colors.length]}">${n}: ${v.toFixed(1)}h (${pct.toFixed(0)}%)</div>`; }).join("");
}

function renderHistory(rows){
  if(!rows.length){ historyBody.innerHTML='<tr><td colspan="7">No history for current filter.</td></tr>'; return; }
  historyBody.innerHTML = rows.map(r=>`<tr><td>${r.changed_at}</td><td>${r.action}</td><td>${r.changed_by}</td><td>${r.employee_name||""}</td><td>${r.project_name||""}</td><td>${r.work_date||""}</td><td>${Number(r.hours_spent||0).toFixed(2)}</td></tr>`).join("");
}

(async()=>{ await initFilters(); await loadDashboard(); })();
</script>
</body>
</html>
"""
