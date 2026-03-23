from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional
import openpyxl
import io
import re
import json
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

GATE_ORDER = ["PCI", "FeG", "CSG", "CG", "DG", "FDG", "FIG", "RG", "EG"]
DATA_FILE = "qtp_data.json"


# ---- JSON DB helpers ----

def read_db():
    if not os.path.exists(DATA_FILE):
        return {"tasks": [], "engineers": [], "project_gates": {}}
    with open(DATA_FILE, "r") as f:
        return json.load(f)

def write_db(db):
    with open(DATA_FILE, "w") as f:
        json.dump(db, f, indent=2)


# ---- Models ----

class Task(BaseModel):
    project_id: str
    item: str
    quality_task: str
    gate: str
    quality_engineer: str
    responsible_engineer: str
    expected_status: str
    actual_status: str
    rating: str
    followup_week: str
    planned_start: Optional[str] = None
    planned_end: Optional[str] = None
    completion_level: str
    comments: Optional[str] = ""
    next_gate: Optional[str] = ""
    deviation: Optional[str] = "No"
    deviation_text: Optional[str] = ""


class DeleteTask(BaseModel):
    project_id: str
    item: str
    quality_task: str
    gate: str


class EngineerName(BaseModel):
    name: str
    email: Optional[str] = ""


class AdvanceGate(BaseModel):
    project_id: str
    new_gate: str


# ---- Engineer helpers ----
def _eng_name(e):
    return e["name"] if isinstance(e, dict) else e

def _eng_email(e):
    return e.get("email", "") if isinstance(e, dict) else ""

def _find_eng(engineers, name):
    for e in engineers:
        if _eng_name(e) == name:
            return e
    return None

def _migrate_engineers(engineers):
    """Migrate plain strings to objects if needed."""
    return [{"name": e, "email": ""} if isinstance(e, str) else e for e in engineers]


# ---------- SAVE TASK ----------
@app.post("/save-task")
def save_task(task: Task):
    db = read_db()
    tasks = db["tasks"]

    # Check if ANY rows exist for this project + item + quality_task
    existing = [t for t in tasks if t["project_id"] == task.project_id
                and t["item"] == task.item and t["quality_task"] == task.quality_task]

    if not existing:
        # FIRST TIME: Insert 9 rows, one per gate
        for g in GATE_ORDER:
            if g == task.gate:
                tasks.append({
                    "project_id": task.project_id, "item": task.item,
                    "quality_task": task.quality_task, "gate": g,
                    "quality_engineer": task.quality_engineer,
                    "responsible_engineer": task.responsible_engineer,
                    "expected_status": task.expected_status,
                    "actual_status": task.actual_status,
                    "rating": task.rating, "followup_week": task.followup_week,
                    "planned_start": task.planned_start or "",
                    "planned_end": task.planned_end or "",
                    "completion_level": task.completion_level,
                    "comments": task.comments or "",
                    "next_gate": task.next_gate or "",
                    "deviation": task.deviation or "No",
                    "deviation_text": task.deviation_text or "",
                    "current_gate": db["project_gates"].get(task.project_id, "PCI")
                })
            else:
                tasks.append({
                    "project_id": task.project_id, "item": task.item,
                    "quality_task": task.quality_task, "gate": g,
                    "quality_engineer": task.quality_engineer,
                    "responsible_engineer": task.responsible_engineer,
                    "expected_status": "", "actual_status": "", "rating": "",
                    "followup_week": "", "planned_start": "", "planned_end": "",
                    "completion_level": "", "comments": "",
                    "next_gate": task.next_gate or "",
                    "deviation": "No", "deviation_text": "",
                    "current_gate": db["project_gates"].get(task.project_id, "PCI")
                })
    else:
        # UPDATE: find the specific gate row and update it
        for t in tasks:
            if (t["project_id"] == task.project_id and t["item"] == task.item
                    and t["quality_task"] == task.quality_task and t["gate"] == task.gate):
                t.update({
                    "quality_engineer": task.quality_engineer,
                    "responsible_engineer": task.responsible_engineer,
                    "expected_status": task.expected_status,
                    "actual_status": task.actual_status,
                    "rating": task.rating,
                    "followup_week": task.followup_week,
                    "planned_start": task.planned_start or "",
                    "planned_end": task.planned_end or "",
                    "completion_level": task.completion_level,
                    "comments": task.comments or "",
                    "next_gate": task.next_gate or "",
                    "deviation": task.deviation or "No",
                    "deviation_text": task.deviation_text or ""
                })
        # Update shared fields across all gates for this task
        for t in tasks:
            if (t["project_id"] == task.project_id and t["item"] == task.item
                    and t["quality_task"] == task.quality_task):
                t["quality_engineer"] = task.quality_engineer
                t["responsible_engineer"] = task.responsible_engineer
                t["next_gate"] = task.next_gate or ""

    write_db(db)
    return {"status": "saved"}


# ---------- LOAD PROJECT ----------
@app.get("/project/{project_id}")
def get_project(project_id: str, gate: Optional[str] = None):
    db = read_db()
    tasks = db["tasks"]

    if gate:
        data = [t for t in tasks if t["project_id"] == project_id and t["gate"] == gate]
    else:
        data = [t for t in tasks if t["project_id"] == project_id]

    current_gate = db["project_gates"].get(project_id, "PCI")
    return {"data": data, "current_gate": current_gate}


# ---------- DELETE TASK ----------
@app.post("/delete-task")
def delete_task(req: DeleteTask):
    db = read_db()
    db["tasks"] = [t for t in db["tasks"] if not (
        t["project_id"] == req.project_id and t["item"] == req.item
        and t["quality_task"] == req.quality_task and t["gate"] == req.gate
    )]
    write_db(db)
    return {"status": "deleted"}


# ---------- ENGINEERS ----------
@app.get("/engineers")
def get_engineers():
    db = read_db()
    engineers = _migrate_engineers(db.get("engineers", []))
    db["engineers"] = engineers
    write_db(db)
    return {"engineers": sorted(engineers, key=lambda e: _eng_name(e).lower())}


@app.post("/save-engineer")
def save_engineer(eng: EngineerName):
    db = read_db()
    db["engineers"] = _migrate_engineers(db.get("engineers", []))
    existing = _find_eng(db["engineers"], eng.name)
    if existing:
        # Update email if provided and not already set
        if eng.email and not existing.get("email"):
            existing["email"] = eng.email
    else:
        db["engineers"].append({"name": eng.name, "email": eng.email or ""})
    write_db(db)
    return {"status": "saved"}


@app.post("/update-engineer-email")
def update_engineer_email(eng: EngineerName):
    db = read_db()
    db["engineers"] = _migrate_engineers(db.get("engineers", []))
    existing = _find_eng(db["engineers"], eng.name)
    if existing and eng.email:
        existing["email"] = eng.email
        write_db(db)
    return {"status": "saved"}


# ---------- ADVANCE GATE ----------
@app.post("/advance-gate")
def advance_gate(req: AdvanceGate):
    db = read_db()
    db["project_gates"][req.project_id] = req.new_gate
    for t in db["tasks"]:
        if t["project_id"] == req.project_id:
            t["current_gate"] = req.new_gate
    write_db(db)
    return {"status": "advanced", "current_gate": req.new_gate}


# ---------- UPLOAD QTP EXCEL ----------
COL = {
    "item": 4, "quality_task": 5, "responsible_engineer": 7,
    "quality_engineer": 8, "expected_status": 9, "completion_level": 10,
    "actual_status": 11, "rating": 12, "deviation": 13,
    "followup_week": 14, "planned_start": 15, "planned_end": 16, "comments": 18,
}
HEADER_ROW = 7
DATA_START = 8


def _clean(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("none", "nan"):
        return ""
    return s


def _parse_date(val):
    if val is None:
        return None
    from datetime import datetime, date
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s or s.lower() in ("none", "nan", "tbd", "n/a", "na", "-"):
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def _find_qtp_sheet(wb):
    for name in wb.sheetnames:
        if "qtp" in name.lower():
            return wb[name]
    return wb.worksheets[0]


@app.post("/upload-qtp")
async def upload_qtp(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
        ws = _find_qtp_sheet(wb)

        project_id = _clean(ws.cell(2, 3).value) or "UNKNOWN"
        project_id = project_id.replace("/", "-").replace("  ", " ").strip()

        current_gate = _clean(ws.cell(2, 13).value) or "PCI"
        current_gate = current_gate.strip()
        if current_gate not in GATE_ORDER:
            for g in GATE_ORDER:
                if g.lower() == current_gate.lower():
                    current_gate = g
                    break

        parsed_tasks = []
        last_item = ""
        last_resp = ""
        last_qe = ""

        for r in range(DATA_START, ws.max_row + 1):
            qt = _clean(ws.cell(r, COL["quality_task"]).value)
            if not qt:
                continue
            item_val = _clean(ws.cell(r, COL["item"]).value)
            if item_val:
                last_item = item_val
            resp_val = _clean(ws.cell(r, COL["responsible_engineer"]).value)
            if resp_val:
                last_resp = resp_val
            qe_val = _clean(ws.cell(r, COL["quality_engineer"]).value)
            if qe_val:
                last_qe = qe_val

            dev = _clean(ws.cell(r, COL["deviation"]).value)
            parsed_tasks.append({
                "item": last_item or "Default",
                "quality_task": qt,
                "responsible_engineer": last_resp.replace("\n", ", "),
                "quality_engineer": last_qe.replace("\n", ", "),
                "expected_status": _clean(ws.cell(r, COL["expected_status"]).value),
                "actual_status": _clean(ws.cell(r, COL["actual_status"]).value),
                "rating": _clean(ws.cell(r, COL["rating"]).value),
                "completion_level": _clean(ws.cell(r, COL["completion_level"]).value),
                "deviation": "No",
                "deviation_text": dev if dev.lower() not in ("no", "na", "") else "",
                "followup_week": _clean(ws.cell(r, COL["followup_week"]).value),
                "planned_start": _parse_date(ws.cell(r, COL["planned_start"]).value),
                "planned_end": _parse_date(ws.cell(r, COL["planned_end"]).value),
                "comments": _clean(ws.cell(r, COL["comments"]).value),
            })

        if not parsed_tasks:
            return {"status": "error", "detail": "No task rows found in the QTP sheet."}

        db = read_db()
        db["project_gates"][project_id] = current_gate

        # Collect engineers
        db["engineers"] = _migrate_engineers(db.get("engineers", []))
        existing_names = {_eng_name(e) for e in db["engineers"]}
        for t in parsed_tasks:
            for name in re.split(r"[,\n]", t["responsible_engineer"]):
                name = name.strip()
                if name and name not in existing_names:
                    db["engineers"].append({"name": name, "email": ""})
                    existing_names.add(name)
            for name in re.split(r"[,\n]", t["quality_engineer"]):
                name = name.strip()
                if name and name not in existing_names:
                    db["engineers"].append({"name": name, "email": ""})
                    existing_names.add(name)

        inserted = 0
        for t in parsed_tasks:
            existing = [x for x in db["tasks"] if x["project_id"] == project_id
                        and x["item"] == t["item"] and x["quality_task"] == t["quality_task"]]
            if existing:
                for x in db["tasks"]:
                    if (x["project_id"] == project_id and x["item"] == t["item"]
                            and x["quality_task"] == t["quality_task"] and x["gate"] == current_gate):
                        x.update({
                            "quality_engineer": t["quality_engineer"],
                            "responsible_engineer": t["responsible_engineer"],
                            "expected_status": t["expected_status"],
                            "actual_status": t["actual_status"],
                            "rating": t["rating"],
                            "followup_week": t["followup_week"],
                            "planned_start": t["planned_start"] or "",
                            "planned_end": t["planned_end"] or "",
                            "completion_level": t["completion_level"],
                            "comments": t["comments"],
                            "deviation": t["deviation"],
                            "deviation_text": t["deviation_text"],
                            "current_gate": current_gate
                        })
            else:
                for g in GATE_ORDER:
                    if g == current_gate:
                        db["tasks"].append({
                            "project_id": project_id, "item": t["item"],
                            "quality_task": t["quality_task"], "gate": g,
                            "quality_engineer": t["quality_engineer"],
                            "responsible_engineer": t["responsible_engineer"],
                            "expected_status": t["expected_status"],
                            "actual_status": t["actual_status"],
                            "rating": t["rating"],
                            "followup_week": t["followup_week"],
                            "planned_start": t["planned_start"] or "",
                            "planned_end": t["planned_end"] or "",
                            "completion_level": t["completion_level"],
                            "comments": t["comments"],
                            "next_gate": "", "deviation": t["deviation"],
                            "deviation_text": t["deviation_text"],
                            "current_gate": current_gate
                        })
                    else:
                        db["tasks"].append({
                            "project_id": project_id, "item": t["item"],
                            "quality_task": t["quality_task"], "gate": g,
                            "quality_engineer": t["quality_engineer"],
                            "responsible_engineer": t["responsible_engineer"],
                            "expected_status": "", "actual_status": "", "rating": "",
                            "followup_week": "", "planned_start": "", "planned_end": "",
                            "completion_level": "", "comments": "", "next_gate": "",
                            "deviation": "No", "deviation_text": "",
                            "current_gate": current_gate
                        })
                inserted += 1

        write_db(db)
        return {
            "status": "ok",
            "project_id": project_id,
            "current_gate": current_gate,
            "tasks_imported": len(parsed_tasks),
            "tasks_new": inserted,
            "tasks_updated": len(parsed_tasks) - inserted,
            "items": list(set(t["item"] for t in parsed_tasks))
        }

    except Exception as e:
        return {"status": "error", "detail": str(e)}


# ---------- ALL PROJECTS (landing page) ----------
@app.get("/all-projects")
def all_projects():
    db = read_db()
    tasks = db["tasks"]
    project_gates = db["project_gates"]
    projects = {}
    for t in tasks:
        pid = t["project_id"]
        if pid not in projects:
            projects[pid] = {"project_id": pid, "tasks": []}
        projects[pid]["tasks"].append(t)

    result = []
    for pid, p in projects.items():
        current_gate = project_gates.get(pid, "PCI")
        gate_tasks = [t for t in p["tasks"] if t["gate"] == current_gate]
        unique_tasks = {t["quality_task"] for t in p["tasks"]}
        total = len(unique_tasks)

        green = yellow = red = deviated = completed = 0
        for t in gate_tasks:
            r = t.get("rating", "")
            a = t.get("actual_status", "")
            if r == "Green": green += 1
            elif r == "Yellow": yellow += 1
            elif r == "Red": red += 1
            if r not in ("Green", "Yellow", "Red") and (r or a or t.get("expected_status", "")):
                deviated += 1
            if a.lower() == "completed":
                completed += 1

        pct = round(completed / total * 100) if total else 0
        result.append({
            "project_id": pid,
            "current_gate": current_gate,
            "total_tasks": total,
            "completed": completed,
            "completion_pct": pct,
            "green": green, "yellow": yellow, "red": red,
            "deviations": deviated
        })

    return {"projects": sorted(result, key=lambda x: x["project_id"])}


# ---------- OVERDUE TASKS ----------
@app.get("/overdue-tasks")
def overdue_tasks():
    from datetime import date
    db = read_db()
    engineers = {_eng_name(e): _eng_email(e) for e in _migrate_engineers(db.get("engineers", []))}
    today = date.today().isoformat()
    overdue = []
    seen = set()
    for t in db["tasks"]:
        end = t.get("planned_end", "")
        status = (t.get("actual_status", "") or "").lower()
        if not end or status == "completed": continue
        if end < today:
            key = (t["project_id"], t["item"], t["quality_task"], t["gate"])
            if key in seen: continue
            seen.add(key)
            resp = t.get("responsible_engineer", "")
            qe = t.get("quality_engineer", "")
            overdue.append({
                "project_id": t["project_id"], "item": t["item"],
                "task": t["quality_task"], "gate": t["gate"],
                "planned_end": end, "actual_status": t.get("actual_status", ""),
                "responsible_engineer": resp,
                "responsible_email": engineers.get(resp, ""),
                "quality_engineer": qe,
                "quality_email": engineers.get(qe, "")
            })
    return {"overdue": overdue}


# ---------- SERVE FRONTEND ----------
app.mount("/", StaticFiles(directory="static", html=True), name="static")

