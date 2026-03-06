import json
import os
from collections import defaultdict
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

from flask import Flask, Response, jsonify, request
from openpyxl import load_workbook

from audit_utils import default_operator
from data_backend import (
    create_entry,
    delete_entry,
    fetch_all_entries,
    fetch_entries,
    fetch_filters,
    fetch_history_rows,
    fetch_monthly_summary,
    import_entries,
    init_backend,
    update_entry,
)
from export_utils import export_excel_report_from_data


ALL_PROJECTS = "All Projects"
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = (Path("/tmp") / "worklog.db") if os.environ.get("VERCEL") else (BASE_DIR / "worklog.db")


app = Flask(__name__)
init_backend(DB_PATH)


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


def _normalize_header(value: object) -> str:
    return str(value or "").strip().lower().replace(" ", "_")


def _parse_import_rows(file_bytes: bytes) -> tuple[list[dict], str]:
    workbook = load_workbook(filename=BytesIO(file_bytes), data_only=True, read_only=True)
    sheet = workbook["All Entries"] if "All Entries" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
    rows = sheet.iter_rows(values_only=True)
    try:
        headers = next(rows)
    except StopIteration:
        return [], "Excel file is empty."

    header_map = {_normalize_header(name): idx for idx, name in enumerate(headers)}
    aliases = {
        "work_date": ["work_date", "date"],
        "employee_name": ["employee_name", "employee"],
        "project_name": ["project_name", "project"],
        "task_name": ["task_name", "task"],
        "hours_spent": ["hours_spent", "hours"],
    }
    col_idx: dict[str, int] = {}
    for key, keys in aliases.items():
        matched = next((header_map[k] for k in keys if k in header_map), None)
        if matched is None:
            return [], f"Missing required column: {key}"
        col_idx[key] = matched

    parsed: list[dict] = []
    errors: list[str] = []
    for row_number, row in enumerate(rows, start=2):
        raw = {k: row[v] if v < len(row) else None for k, v in col_idx.items()}
        if not any(value not in (None, "") for value in raw.values()):
            continue
        work_date_value = raw["work_date"]
        if isinstance(work_date_value, datetime):
            work_date = work_date_value.strftime("%Y-%m-%d")
        elif isinstance(work_date_value, date):
            work_date = work_date_value.isoformat()
        else:
            work_date = str(work_date_value or "").strip()
        employee_name = str(raw["employee_name"] or "").strip()
        project_name = str(raw["project_name"] or "").strip()
        task_name = str(raw["task_name"] or "").strip()
        hours_raw = raw["hours_spent"]

        try:
            datetime.strptime(work_date, "%Y-%m-%d")
            hours = round(float(hours_raw), 2)
            if hours <= 0 or not employee_name or not project_name or not task_name:
                raise ValueError
        except (TypeError, ValueError):
            errors.append(f"Invalid data at row {row_number}.")
            continue

        parsed.append(
            {
                "employee_name": employee_name,
                "work_date": work_date,
                "project_name": project_name,
                "task_name": task_name,
                "hours_spent": hours,
            }
        )

    if errors:
        return [], " ".join(errors[:5])
    if not parsed:
        return [], "No valid rows found in Excel file."
    return parsed, ""


@app.get("/")
def index() -> Response:
    html = HTML_PAGE.replace("__DEFAULT_OPERATOR__", json.dumps(default_operator())[1:-1])
    return Response(html, mimetype="text/html; charset=utf-8")


@app.get("/api/filters")
def api_filters():
    return jsonify(fetch_filters(DB_PATH))


@app.get("/api/dashboard")
def api_dashboard():
    month = request.args.get("month") or date.today().strftime("%Y-%m")
    project = request.args.get("project") or ALL_PROJECTS
    entries = fetch_entries(DB_PATH, month, project)
    history = fetch_history_rows(DB_PATH, month=month, project=project, limit=40)
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

    create_entry(DB_PATH, values, operator)
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

    if not update_entry(DB_PATH, entry_id, values, operator):
        return jsonify({"error": "Entry not found."}), 404
    return jsonify({"status": "ok"})


@app.delete("/api/entries/<int:entry_id>")
def api_delete_entry(entry_id: int):
    payload = request.get_json(silent=True) or {}
    operator = str(payload.get("operator_name", "")).strip()
    if not operator:
        return jsonify({"error": "Please enter the operator name."}), 400

    if not delete_entry(DB_PATH, entry_id, operator):
        return jsonify({"error": "Entry not found."}), 404
    return jsonify({"status": "ok"})


@app.post("/api/import.xlsx")
def api_import_excel():
    operator = str(request.form.get("operator_name", "")).strip()
    if not operator:
        return jsonify({"error": "Please enter the operator name."}), 400
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "Please choose an Excel file (.xlsx)."}), 400
    if not upload.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are supported."}), 400
    try:
        entries, error = _parse_import_rows(upload.read())
    except Exception:
        return jsonify({"error": "Could not read Excel file."}), 400
    if error:
        return jsonify({"error": error}), 400

    overwrite = str(request.form.get("overwrite", "1")).strip().lower() in {"1", "true", "yes", "on"}
    result = import_entries(DB_PATH, entries, operator, overwrite=overwrite)
    return jsonify({"status": "ok", "result": result})


@app.get("/api/export.xlsx")
def api_export_excel():
    month = request.args.get("month") or date.today().strftime("%Y-%m")
    project = request.args.get("project") or ALL_PROJECTS
    all_entries = fetch_all_entries(DB_PATH)
    filtered_entries = fetch_entries(DB_PATH, month, project)
    monthly_summary = fetch_monthly_summary(DB_PATH)
    file_bytes = export_excel_report_from_data(
        all_entries=all_entries,
        filtered_entries=filtered_entries,
        monthly_summary=monthly_summary,
        month=month,
        project=project,
    )
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
        <div class="field" style="margin-top:8px;">
          <label>Import Excel (.xlsx)</label>
          <input id="importFile" type="file" accept=".xlsx" />
        </div>
        <label style="display:flex; align-items:center; gap:8px; margin-top:6px;">
          <input id="overwriteImport" type="checkbox" checked style="width:auto; padding:0;" />
          Overwrite duplicate records (same Date + Employee + Project + Task)
        </label>
        <button id="importButton" type="button" style="margin-top:8px; background:#4d7d5b;">Import Excel</button>
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
document.getElementById("importButton").addEventListener("click", async ()=>{
  const op=document.getElementById("operator_name").value.trim();
  const fileInput=document.getElementById("importFile");
  const file=fileInput.files && fileInput.files[0];
  if(!op){ formStatus.textContent="Please enter the operator name."; formStatus.className="status error"; return; }
  if(!file){ formStatus.textContent="Please choose an Excel file to import."; formStatus.className="status error"; return; }
  const overwrite=document.getElementById("overwriteImport").checked;
  if(!confirm(overwrite ? "Import file and overwrite duplicate records?" : "Import file without overwriting duplicates?")) return;
  const fd=new FormData();
  fd.append("file", file);
  fd.append("operator_name", op);
  fd.append("overwrite", overwrite ? "1" : "0");
  const resp=await fetch("/api/import.xlsx",{method:"POST",body:fd});
  const res=await resp.json();
  if(!resp.ok){ formStatus.textContent=res.error||"Import failed."; formStatus.className="status error"; return; }
  const r=res.result||{};
  formStatus.textContent=`Import done. Created: ${r.created||0}, Updated: ${r.updated||0}, Skipped: ${r.skipped||0}.`;
  formStatus.className="status";
  fileInput.value="";
  await initFilters();
  await loadDashboard();
});
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
