import hashlib
import re
from datetime import UTC, date, datetime, timedelta
from typing import Annotated, Literal
from urllib.parse import urlsplit

from fastapi import APIRouter, Depends, Header, Query
from pydantic import BaseModel, Field, field_validator

from .identity import fail, user_for
from .store import DSN, ZONE, audit, database, instant, local_today, pack, uid, unpack

router = APIRouter(prefix="/api/v1")
User = Annotated[dict, Depends(user_for)]
Idem = Annotated[str, Header(alias="Idempotency-Key", max_length=128)]
ACTIVE = ("pending", "in_progress", "paused", "returned")
CLOSED = ("approved", "revoked", "closed")
ALLOWED_SORTED = set(ACTIVE + CLOSED + ("submitted", "scheduled"))
FILE_COLUMNS = "id, owner_id, assignment_id, template_id, parent_id, run_id, name, extension, digest, size, label, created_at"

BASE = """SELECT a.*, p.name AS assignee_name, p.department, o.name AS owner_name,
          r.name AS reviewer_name FROM assignments a JOIN users p ON p.id=a.assignee_id
          JOIN users o ON o.id=a.owner_id LEFT JOIN users r ON r.id=a.reviewer_id"""


def now_or_none(value):
    if value and value.utcoffset() is None:
        raise ValueError("时间必须包含时区")
    return value


class TaskFields(BaseModel):
    title: str = Field(min_length=1, max_length=130)
    description: str = Field(default="", max_length=8000)
    deliverable: str = Field(default="", max_length=2000)
    kind: Literal["docx", "pptx", "other"] = "docx"
    priority: Literal["normal", "high", "urgent"] = "normal"
    work_date: date
    due_at: datetime | None = None
    assignee_ids: list[str] = Field(min_length=1, max_length=100)
    reviewer_id: str | None = None
    publish_at: datetime | None = None
    publish_mode: Literal["draft", "immediate"] = "draft"
    repeat: Literal["once", "weekdays"] = "once"
    end_date: date | None = None
    exclude_dates: list[date] = Field(default_factory=list, max_length=180)
    include_dates: list[date] = Field(default_factory=list, max_length=180)

    @field_validator("due_at", "publish_at")
    @classmethod
    def validate_timezone(cls, value): 
        return now_or_none(value)

    @field_validator("title", "description", "deliverable")
    @classmethod
    def trim(cls, value):
        return value.strip()

    def validate_dates(self):
        if self.repeat == "weekdays" and (not self.end_date or not 0 <= (self.end_date - self.work_date).days <= 366):
            fail(422, "工作日计划需填写结束日期，最长 366 天")
        if self.due_at and not date(2000, 1, 1) <= self.due_at.date() <= date(2099, 12, 31):
            fail(422, "截止日期超出有效范围")
        if self.publish_at and self.due_at and self.publish_at > self.due_at:
            fail(422, "发布时间须早于截止时间")
        if not self.title or not date(2000, 1, 1) <= self.work_date <= date(2099, 12, 30):
            fail(422, "请填写有效的任务标题与工作日期")


class TemplateChange(TaskFields):
    version: int = Field(ge=1)


class PublishBody(BaseModel):
    version: int = Field(ge=1)


class EditAssignment(BaseModel):
    version: int = Field(ge=1)
    title: str | None = Field(default=None, min_length=1, max_length=130)
    description: str | None = Field(default=None, max_length=8000)
    deliverable: str | None = Field(default=None, max_length=2000)
    due_at: datetime | None = None
    priority: Literal["normal", "high", "urgent"] | None = None
    assignee_id: str | None = None
    reviewer_id: str | None = None

    @field_validator("due_at")
    @classmethod
    def validate_timezone(cls, value):
        return now_or_none(value)


class ActionBody(BaseModel):
    version: int = Field(ge=1)
    note: str = Field(default="", max_length=3000)


class SubmitBody(BaseModel):
    version: int = Field(ge=1)
    note: str = Field(default="", max_length=4000)
    artifact_ids: list[str] = Field(default_factory=list, max_length=12)
    link: str = Field(default="", max_length=500)


class ReviewBody(BaseModel):
    version: int = Field(ge=1)
    decision: Literal["approve", "return"]
    note: str = Field(default="", max_length=3000)


def is_manager(user):
    return user["role"] in {"admin", "manager"}


def require_manager(user):
    if not is_manager(user):
        fail(403, "需要主管或管理员权限")


def manager_scope(user, department):
    return user["role"] == "admin" or user["role"] == "manager" and user["department"] == department


def check_people(db, user, assignees, reviewer=None):
    for user_id in dict.fromkeys(assignees + ([reviewer] if reviewer else [])):
        person = db.one("SELECT * FROM users WHERE id=?", (user_id,))
        if not person or not person["enabled"] or not manager_scope(user, person["department"]):
            fail(403, "所选人员不在可分派范围内")
        if user_id in assignees and person["role"] not in {"employee", "manager"}:
            fail(422, "请选择可办理任务的员工")
        if reviewer and user_id == reviewer and (person["role"] not in {"reviewer", "manager", "admin"} or user_id in assignees):
            fail(422, "请选择独立的审核人")


def template_for(db, user, key):
    row = db.one("SELECT * FROM templates WHERE id=?", (key,))
    if not row or not (user["role"] == "admin" or user["role"] == "manager" and row["owner_id"] == user["id"]):
        fail(404, "任务草稿或计划不存在")
    return row


def sees(user, task):
    if user["id"] == task["assignee_id"]:
        return task["status"] != "scheduled" or manager_scope(user, task["department"])
    return manager_scope(user, task["department"]) or user["id"] == task["reviewer_id"] and task["status"] != "scheduled"


def task_for(db, key, user, actor="read"):
    row = db.one(f"{BASE} WHERE a.id=?", (key,))
    if not row or not sees(user, row):
        fail(404, "任务不存在或当前账号无权访问")
    if actor == "employee" and (row["assignee_id"] != user["id"] or row["status"] == "scheduled"):
        fail(403, "此操作需要经办员工")
    if actor == "manager" and not manager_scope(user, row["department"]):
        fail(403, "此操作需要任务管理权限")
    if actor == "reviewer" and not (manager_scope(user, row["department"]) or row["reviewer_id"] == user["id"]):
        fail(403, "此操作需要审核权限")
    if actor == "reviewer" and user["id"] == row["assignee_id"]:
        fail(403, "经办人无法审核自己的任务")
    return row


def version_of(row, expected):
    if row["version"] != expected:
        fail(409, "任务已被更新，请刷新后重试。当前输入可以保留")


def change_status(db, row, status):
    db.execute("UPDATE assignments SET status=?, version=version+1, updated_at=? WHERE id=?",
               (status, instant(), row["id"]))


def as_task(row):
    item = dict(row)
    item["overdue"] = bool(item["due_at"] and item["due_at"] < instant() and item["status"] not in CLOSED + ("scheduled",))
    return item


def get_saved(db, user, action, key, payload):
    if not re.fullmatch(r"[A-Za-z0-9_.:-]{8,128}", key):
        fail(422, "请求幂等键无效")
    fingerprint = hashlib.sha256(pack(payload).encode()).hexdigest()
    row = db.one("SELECT fingerprint,result FROM idempotency WHERE actor_id=? AND action=? AND key=?",
                 (user["id"], action, key))
    if row and row["fingerprint"] != fingerprint:
        fail(409, "请求已用于不同内容，请刷新操作")
    return (unpack(row["result"]) if row else None), fingerprint


def save_result(db, user, action, key, fingerprint, result):
    db.execute("INSERT INTO idempotency(actor_id,action,key,fingerprint,result,created_at) VALUES(?,?,?,?,?,?)",
               (user["id"], action, key, fingerprint, pack(result), instant()))
    return result


def date_allowed(spec, day):
    text = day.isoformat()
    return text not in spec.get("exclude_dates", []) and (day.weekday() < 5 or text in spec.get("include_dates", []))


def time_for(spec, field, day):
    value = spec.get(field)
    if not value:
        return ""
    moment = datetime.fromisoformat(value).astimezone(ZONE)
    original_day = date.fromisoformat(spec["work_date"])
    relative = moment.date() - original_day
    target_day = day + relative
    return instant(datetime.combine(target_day, moment.timetz()))


def materialize(db, key, owner, spec, day):
    now = instant()
    publish = time_for(spec, "publish_at", day) or now
    due = time_for(spec, "due_at", day)
    status = "scheduled" if publish > now else "pending"
    ids = []
    for person in dict.fromkeys(spec["assignee_ids"]):
        existing = db.one("SELECT id FROM assignments WHERE template_id=? AND assignee_id=? AND work_date=?",
                          (key, person, day.isoformat()))
        if existing:
            ids.append(existing["id"])
            continue
        assignment_id = uid()
        db.execute("""INSERT INTO assignments(id,template_id,assignee_id,owner_id,reviewer_id,title,description,work_date,
                   due_at,priority,kind,deliverable,status,publish_at,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                   (assignment_id, key, person, owner, spec["reviewer_id"], spec["title"], spec["description"],
                    day.isoformat(), due, spec["priority"], spec["kind"], spec["deliverable"], status, publish, now, now))
        audit(db, "task.scheduled" if status == "scheduled" else "task.published", owner, assignment_id, key, assignee_id=person)
        ids.append(assignment_id)
    return ids


def release(db, row):
    spec = unpack(row["data"])
    day = date.fromisoformat(spec["work_date"])
    if spec["repeat"] == "weekdays":
        state, next_day = "active", max(day, local_today()).isoformat()
        db.execute("UPDATE templates SET state=?, next_date=?, version=version+1 WHERE id=?", (state, next_day, row["id"]))
        audit(db, "schedule.created", row["owner_id"], template=row["id"], work_date=day.isoformat())
        return []
    db.execute("UPDATE templates SET state='published', version=version+1 WHERE id=?", (row["id"],))
    return materialize(db, row["id"], row["owner_id"], spec, day)


def tick():
    today = local_today()
    now = instant()
    with database() as db:
        due = db.one("SELECT id FROM assignments WHERE status='scheduled' AND publish_at<=? LIMIT 1", (now,))
        schedule = db.one("SELECT id FROM templates WHERE state='active' AND next_date<=? LIMIT 1", (today.isoformat(),))
    if not due and not schedule:
        return
    with database(write=True) as db:
        for row in db.all("SELECT * FROM templates WHERE state='active' AND next_date<=?", (today.isoformat(),)):
            spec = unpack(row["data"])
            day = date.fromisoformat(row["next_date"])
            last = date.fromisoformat(spec["end_date"])
            owner = db.one("SELECT * FROM users WHERE id=?", (row["owner_id"],))
            while day <= min(today, last):
                if date_allowed(spec, day):
                    allowed = [person["id"] for user_id in spec["assignee_ids"]
                               if (person := db.one("SELECT * FROM users WHERE id=?", (user_id,)))
                               and person["role"] in {"employee", "manager"} and owner["enabled"] and person["enabled"] and manager_scope(owner, person["department"])]
                    if allowed:
                        reviewer = db.one("SELECT * FROM users WHERE id=?", (spec["reviewer_id"],)) if spec["reviewer_id"] else None
                        assigned_reviewer = reviewer["id"] if reviewer and reviewer["enabled"] and reviewer["role"] in {"reviewer", "manager", "admin"} and manager_scope(owner, reviewer["department"]) and reviewer["id"] not in allowed else None
                        materialize(db, row["id"], row["owner_id"], dict(spec, assignee_ids=allowed, reviewer_id=assigned_reviewer), day)
                day += timedelta(days=1)
            db.execute("UPDATE templates SET state=?,next_date=? WHERE id=?",
                       ("finished" if day > last else "active", day.isoformat(), row["id"]))
        for row in db.all("SELECT * FROM assignments WHERE status='scheduled' AND publish_at<=?", (now,)):
            change_status(db, row, "pending")
            audit(db, "task.published", row["owner_id"], row["id"], row["template_id"], assignee_id=row["assignee_id"])


@router.get("/me")
def current(user: User):
    return {**user, "today": local_today().isoformat(), "timezone": str(ZONE)}


@router.get("/people")
def people(user: User):
    require_manager(user)
    with database() as db:
        query = "SELECT id,name,department,role FROM users WHERE enabled=1"
        return [dict(row) for row in db.all(query + " ORDER BY department,name")] if user["role"] == "admin" else [
            dict(row) for row in db.all(query + " AND department=? ORDER BY name", (user["department"],))
        ]


@router.post("/tasks")
def create_task(payload: TaskFields, user: User, key: Idem):
    require_manager(user)
    payload.validate_dates()
    spec = payload.model_dump(mode="json")
    with database(write=True) as db:
        saved, fingerprint = get_saved(db, user, "create", key, spec)
        if saved:
            return saved
        check_people(db, user, payload.assignee_ids, payload.reviewer_id)
        task_id = uid()
        db.execute("INSERT INTO templates(id,owner_id,data,state,created_at) VALUES(?,?,?,?,?)",
                   (task_id, user["id"], pack(spec), "draft", instant()))
        audit(db, "task.created", user["id"], template=task_id)
        ids = release(db, template_for(db, user, task_id)) if payload.publish_mode == "immediate" else []
        response = save_result(db, user, "create", key, fingerprint,
                               {"id": task_id, "assignment_ids": ids, "version": 2 if payload.publish_mode == "immediate" else 1})
    if payload.publish_mode == "immediate" and payload.repeat != "once":
        tick()
    return response


@router.patch("/tasks/{key}")
def edit_template(key: str, payload: TemplateChange, user: User):
    require_manager(user)
    payload.validate_dates()
    with database(write=True) as db:
        row = template_for(db, user, key)
        version_of(row, payload.version)
        if row["state"] not in {"draft", "active"}:
            fail(409, "此计划已结束")
        spec = payload.model_dump(mode="json", exclude={"version"})
        check_people(db, user, payload.assignee_ids, payload.reviewer_id)
        if row["state"] == "active":
            original = unpack(row["data"])
            if payload.repeat != "weekdays" or spec["work_date"] != original["work_date"]:
                fail(422, "已启动计划可以调整未来安排，起始日期和重复方式原样保留")
        db.execute("UPDATE templates SET data=?, version=version+1 WHERE id=?", (pack(spec), key))
        audit(db, "schedule.updated" if row["state"] == "active" else "task.draft_updated", user["id"], template=key)
        return {"id": key, "version": payload.version + 1}


@router.post("/tasks/{key}/publish")
def publish(key: str, body: PublishBody, user: User, idem: Idem):
    require_manager(user)
    action = "publish:" + key
    with database(write=True) as db:
        saved, fingerprint = get_saved(db, user, action, idem, body.model_dump())
        if saved:
            return saved
        row = template_for(db, user, key)
        version_of(row, body.version)
        if row["state"] != "draft":
            fail(409, "任务已发布")
        spec = unpack(row["data"])
        check_people(db, user, spec["assignee_ids"], spec["reviewer_id"])
        ids = release(db, row)
        result = save_result(db, user, action, idem, fingerprint, {"id": key, "assignment_ids": ids, "version": body.version + 1})
    tick()
    return result


@router.get("/admin/plans")
def plans(user: User):
    require_manager(user)
    with database() as db:
        scope, args = ("", ()) if user["role"] == "admin" else (" AND owner_id=?", (user["id"],))
        result = []
        for row in db.all("SELECT * FROM templates WHERE state IN ('draft','active','finished')" + scope +
                          " ORDER BY created_at DESC LIMIT 100", args):
            item = dict(row)
            item["data"] = unpack(row["data"])
            item["files"] = [dict(f) for f in db.all(f"SELECT {FILE_COLUMNS} FROM artifacts WHERE template_id=? ORDER BY created_at", (row["id"],))]
            result.append(item)
        return result


@router.get("/admin/tasks")
def admin_tasks(user: User, department: str = "", status: str = "", work_date: date | None = None,
                offset: int = Query(0, ge=0), limit: int = Query(100, ge=1, le=200)):
    if user["role"] not in {"admin", "manager", "reviewer"}:
        fail(403, "需要管理或审核权限")
    tick()
    query, args = BASE + " WHERE 1=1", []
    if user["role"] == "reviewer":
        query += " AND a.reviewer_id=? AND a.status<>'scheduled'"
        args.append(user["id"])
    if user["role"] == "manager":
        query += " AND p.department=?"
        args.append(user["department"])
    if department:
        query += " AND p.department=?"
        args.append(department)
    if status == "overdue":
        query += " AND a.due_at<>'' AND a.due_at<? AND a.status NOT IN ('approved','closed','revoked','scheduled')"
        args.append(instant())
    elif status in ALLOWED_SORTED:
        query += " AND a.status=?"
        args.append(status)
    if work_date:
        query += " AND a.work_date=?"
        args.append(work_date.isoformat())
    with database() as db:
        rows = db.all(query + " ORDER BY a.work_date DESC,a.created_at DESC,a.id LIMIT ? OFFSET ?", (*args, limit, offset))
        return [as_task(row) for row in rows]


@router.get("/me/today")
def today_tasks(user: User, limit: int = Query(150, ge=1, le=200), offset: int = Query(0, ge=0)):
    tick()
    today = local_today().isoformat()
    with database() as db:
        rows = db.all(BASE + """ WHERE a.assignee_id=? AND a.status<>'scheduled' AND
                        (a.work_date=? OR a.status='returned' OR (a.due_at<>'' AND a.due_at<?
                         AND a.status NOT IN ('approved','closed','revoked'))) ORDER BY
                         CASE WHEN a.status='returned' THEN 0 ELSE 1 END,a.due_at,a.id LIMIT ? OFFSET ?""",
                      (user["id"], today, instant(), limit, offset))
    return {"today": today, "timezone": str(ZONE), "items": [as_task(row) for row in rows]}


@router.get("/me/tasks")
def my_tasks(user: User, cursor: int = Query(0, ge=0), offset: int = Query(0, ge=0), limit: int = Query(100, ge=1, le=200)):
    tick()
    with database() as db:
        if cursor:
            prior = "e.payload::jsonb ->> 'previous_assignee'" if DSN else "json_extract(e.payload,'$.previous_assignee')"
            events = db.all("SELECT e.seq,e.assignment_id,e.type,e.created_at FROM events e "
                            "JOIN assignments a ON a.id=e.assignment_id WHERE e.seq>? AND "
                            "((a.assignee_id=? AND a.status<>'scheduled') OR "
                            f"(e.type='task.updated' AND {prior}=?)) ORDER BY e.seq LIMIT ?",
                            (cursor, user["id"], user["id"], limit))
            items = []
            for key in dict.fromkeys(row["assignment_id"] for row in events):
                row = db.one(BASE + " WHERE a.id=? AND a.assignee_id=? AND a.status<>'scheduled'", (key, user["id"]))
                items.append(as_task(row) if row else {"id": key, "removed": True})
            return {"items": items, "cursor": events[-1]["seq"] if events else cursor}
        latest = db.one("SELECT COALESCE(MAX(seq),0) AS n FROM events")["n"]
        rows = db.all(BASE + " WHERE a.assignee_id=? AND a.status<>'scheduled' ORDER BY a.work_date DESC,a.created_at DESC,a.id LIMIT ? OFFSET ?",
                      (user["id"], limit, offset))
        return {"items": [as_task(row) for row in rows], "cursor": latest}


@router.get("/assignments/{key}")
def detail(key: str, user: User):
    with database() as db:
        task = task_for(db, key, user)
        item = as_task(task)
        item["materials"] = [dict(row) for row in db.all(f"SELECT {FILE_COLUMNS} FROM artifacts WHERE template_id=? AND label='material' ORDER BY created_at",
                                                         (task["template_id"],))]
        item["files"] = [dict(row) for row in db.all(f"SELECT {FILE_COLUMNS} FROM artifacts WHERE assignment_id=? AND owner_id=? ORDER BY created_at DESC",
                                                     (key, user["id"]))]
        events = db.all("SELECT e.*,u.name FROM events e LEFT JOIN users u ON u.id=e.actor_id "
                        "WHERE e.assignment_id=? AND e.type NOT IN ('file.download','file.preview','agent.completed','agent.failed','file.generated','file.uploaded','file.edited') "
                        "ORDER BY e.seq DESC LIMIT 80", (key,))
        item["events"] = [{**dict(row), "payload": unpack(row["payload"])} for row in events]
        submissions = db.all("SELECT * FROM submissions WHERE assignment_id=? ORDER BY revision DESC", (key,))
        item["submissions"] = []
        for row in submissions:
            files = db.all("SELECT a.id,a.name,a.extension,a.size,a.digest,a.parent_id FROM artifacts a "
                           "JOIN submission_files f ON a.id=f.artifact_id WHERE f.submission_id=? ORDER BY a.created_at", (row["id"],))
            item["submissions"].append({**dict(row), "files": [dict(file) for file in files]})
        return item


@router.patch("/assignments/{key}")
def update_assignment(key: str, action: EditAssignment, user: User):
    with database(write=True) as db:
        row = task_for(db, key, user, "manager")
        version_of(row, action.version)
        if row["status"] in CLOSED:
            fail(409, "此任务已经结束")
        changes = action.model_dump(exclude_unset=True, exclude={"version"})
        updates = {}
        for name in ("title", "description", "deliverable", "priority"):
            if name in changes and changes[name] is not None:
                value = changes[name].strip()
                if name == "title" and not value:
                    fail(422, "任务标题不能为空")
                updates[name] = value
        if "due_at" in changes:
            updates["due_at"] = instant(action.due_at) if action.due_at else ""
            if row["status"] == "scheduled" and updates["due_at"] and updates["due_at"] < row["publish_at"]:
                fail(422, "发布时间须早于截止时间")
        if action.assignee_id and action.assignee_id != row["assignee_id"]:
            if row["status"] not in ("scheduled", "pending") or db.one("SELECT id FROM artifacts WHERE assignment_id=? LIMIT 1", (key,)) or db.one("SELECT id FROM runs WHERE assignment_id=? LIMIT 1", (key,)):
                fail(409, "已有办理记录的任务无法转派")
            if db.one("SELECT id FROM assignments WHERE template_id=? AND assignee_id=? AND work_date=?",
                      (row["template_id"], action.assignee_id, row["work_date"])):
                fail(409, "该员工已分配此任务")
            updates["assignee_id"] = action.assignee_id
        if "reviewer_id" in changes:
            updates["reviewer_id"] = action.reviewer_id
        if "reviewer_id" in updates or "assignee_id" in updates:
            check_people(db, user, [updates.get("assignee_id", row["assignee_id"])], updates.get("reviewer_id", row["reviewer_id"]))
        updates = {name: value for name, value in updates.items() if row[name] != value}
        if not updates:
            return as_task(row)
        columns = ",".join(f"{name}=?" for name in updates)
        db.execute(f"UPDATE assignments SET {columns},version=version+1,updated_at=? WHERE id=?",
                   (*updates.values(), instant(), key))
        audit(db, "task.updated", user["id"], key, row["template_id"],
              before={name: row[name] for name in updates}, after=updates,
              previous_assignee=row["assignee_id"] if "assignee_id" in updates else None)
    return detail(key, user)


@router.post("/assignments/{key}/action/{action}")
def task_action(key: str, action: Literal["start", "pause", "comment", "extension", "revoke", "close"], body: ActionBody, user: User):
    transitions = {
        "start": (("pending", "paused", "returned"), "in_progress"),
        "pause": (("in_progress",), "paused"),
        "revoke": (("scheduled", *ACTIVE, "submitted"), "revoked"),
        "close": (ACTIVE + ("submitted",), "closed"),
    }
    actor = "manager" if action in ("revoke", "close") else "read" if action == "comment" else "employee"
    with database(write=True) as db:
        row = task_for(db, key, user, actor)
        version_of(row, body.version)
        note = body.note.strip()
        if action in transitions:
            allowed, new = transitions[action]
            if row["status"] not in allowed:
                fail(409, "任务当前状态无法执行此操作")
            change_status(db, row, new)
            if new in {"closed", "revoked"}:
                db.execute("UPDATE submissions SET status=? WHERE assignment_id=? AND status='submitted'", (new, key))
        elif not note or row["status"] in CLOSED + ("scheduled",) or action == "extension" and row["status"] not in ACTIVE:
            fail(422, "请填写说明并确认任务仍可办理")
        audit(db, f"task.{action}", user["id"], key, row["template_id"], note=note)
    return detail(key, user)


@router.post("/assignments/{key}/submissions")
def submit(key: str, body: SubmitBody, user: User, idem: Idem):
    note, link = body.note.strip(), body.link.strip()
    ids = list(dict.fromkeys(body.artifact_ids))
    if not ids and not link:
        fail(422, "请至少提交一个成果文件或交付链接")
    if link:
        url = urlsplit(link)
        if url.scheme != "https" or not url.hostname or url.username or url.password or any(c.isspace() for c in link):
            fail(422, "交付链接需为有效的 HTTPS 地址")
    action = "submit:" + key
    with database(write=True) as db:
        row = task_for(db, key, user, "employee")
        saved, fingerprint = get_saved(db, user, action, idem, body.model_dump())
        if saved:
            return saved
        version_of(row, body.version)
        if row["status"] not in ACTIVE:
            fail(409, "任务当前无法提交")
        for file_id in ids:
            file = db.one("SELECT * FROM artifacts WHERE id=? AND owner_id=? AND assignment_id=? AND label<>'material'",
                          (file_id, user["id"], key))
            if not file:
                fail(422, "此成果文件未关联到当前员工的任务")
        revision = db.one("SELECT COALESCE(MAX(revision),0)+1 AS next FROM submissions WHERE assignment_id=?", (key,))["next"]
        sub = uid()
        db.execute("""INSERT INTO submissions(id,assignment_id,author_id,revision,status,note,link,created_at)
                      VALUES(?,?,?,?,?,?,?,?)""", (sub, key, user["id"], revision, "submitted", note, link, instant()))
        for file in ids:
            db.execute("INSERT INTO submission_files(submission_id,artifact_id) VALUES(?,?)", (sub, file))
        change_status(db, row, "submitted")
        audit(db, "task.submitted", user["id"], key, row["template_id"], revision=revision, submission_id=sub)
        return save_result(db, user, action, idem, fingerprint, {"id": sub, "revision": revision})


@router.post("/submissions/{key}/review")
def review(key: str, body: ReviewBody, user: User, idem: Idem):
    action = "review:" + key
    with database(write=True) as db:
        sub = db.one("SELECT * FROM submissions WHERE id=?", (key,))
        if not sub:
            fail(404, "提交记录不存在")
        row = task_for(db, sub["assignment_id"], user, "reviewer")
        saved, fingerprint = get_saved(db, user, action, idem, body.model_dump())
        if saved:
            return saved
        version_of(row, body.version)
        if row["status"] != "submitted" or sub["status"] != "submitted":
            fail(409, "提交记录已审核或任务已结束")
        note = body.note.strip()
        if body.decision == "return" and not note:
            fail(422, "退回时请填写修改意见")
        new = "approved" if body.decision == "approve" else "returned"
        db.execute("UPDATE submissions SET status=?,review_note=?,reviewed_by=?,reviewed_at=? WHERE id=?",
                   (new, note, user["id"], instant(), key))
        change_status(db, row, new)
        audit(db, "task." + new, user["id"], row["id"], row["template_id"], note=note, revision=sub["revision"])
        return save_result(db, user, action, idem, fingerprint, {"id": key, "status": new})


@router.get("/admin/audit")
def audit_log(user: User, offset: int = Query(0, ge=0), limit: int = Query(100, ge=1, le=200)):
    if user["role"] != "admin":
        fail(403, "需要超级管理员权限")
    with database() as db:
        events = db.all("SELECT e.*,u.name,a.title FROM events e LEFT JOIN users u ON u.id=e.actor_id "
                        "LEFT JOIN assignments a ON a.id=e.assignment_id ORDER BY e.seq DESC LIMIT ? OFFSET ?", (limit, offset))
        return [{**dict(row), "payload": {key: value for key, value in unpack(row["payload"]).items()
                                         if key not in {"before", "after", "note"}}} for row in events]
