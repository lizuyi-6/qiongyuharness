import asyncio
import os
from typing import Annotated, Any, Literal
from urllib.parse import quote

from fastapi import APIRouter, File, Form, UploadFile
from fastapi.responses import Response
from pydantic import BaseModel, Field

from . import gateway, prompting
from .identity import fail
from .office import (MAX_FILE, MIMES, Plan, add_file, clean_text, edit_copy, extract, powerpoint,
                     safe_name, verify, word)
from .projects import shared_file_access
from .store import audit, database, instant, pack, uid, unpack
from .tasks import ACTIVE, FILE_COLUMNS, User, sees, task_for, template_for

router = APIRouter(prefix="/api/v1")
SCANNER = os.getenv("QO_CLAMD", "")
PLAN_SLOTS = asyncio.Semaphore(4)


def file_for(db, user, key):
    row = db.one("SELECT * FROM artifacts WHERE id=?", (key,))
    if not row:
        fail(404, "文件不存在")
    if row["label"] == "material":
        owner = db.one("SELECT owner_id FROM templates WHERE id=?", (row["template_id"],))
        if owner and (owner["owner_id"] == user["id"] or user["role"] == "admin"):
            return row
        related = db.all("SELECT a.*,u.department FROM assignments a JOIN users u ON u.id=a.assignee_id "
                         "WHERE a.template_id=?", (row["template_id"],))
        if any(sees(user, item) for item in related):
            return row
    elif row["owner_id"] == user["id"]:
        if row["assignment_id"]:
            task_for(db, row["assignment_id"], user, "employee")
        return row
    elif row["assignment_id"] and db.one("SELECT 1 FROM submission_files WHERE artifact_id=?", (key,)):
        task_for(db, row["assignment_id"], user, "reviewer")
        return row
    elif not row["assignment_id"] and shared_file_access(db,key,user["id"]):
        return row
    fail(404, "文件不存在或当前账号无权访问")


def public_file(db, key):
    return dict(db.one(f"SELECT {FILE_COLUMNS} FROM artifacts WHERE id=?", (key,)))


def writable_task(db, key, user):
    if not key:
        return None
    row = task_for(db, key, user, "employee")
    if row["status"] not in ACTIVE:
        fail(409, "任务当前状态无法添加文件或启动协助")
    return row


def file_type(row):
    return bytes(row["data"]), row["extension"]


async def upload_data(file):
    try:
        return await file.read(MAX_FILE + 1)
    finally:
        await file.close()


@router.get("/files")
def own_files(user: User):
    with database() as db:
        return [dict(row) for row in db.all(
            f"SELECT {FILE_COLUMNS} FROM artifacts WHERE owner_id=? AND label<>'material' ORDER BY created_at DESC LIMIT 120",
            (user["id"],),
        )]


@router.post("/tasks/{key}/materials")
async def add_material(key: str, user: User, file: Annotated[UploadFile, File()]):
    with database() as db:
        if template_for(db, user, key)["state"] != "draft":
            fail(409, "请在发布前添加任务资料")
    data = await upload_data(file)
    name, ext = await asyncio.to_thread(verify, file.filename or "", data, SCANNER)
    with database(write=True) as db:
        if template_for(db, user, key)["state"] != "draft":
            fail(409, "任务已经发布")
        count = db.one("SELECT COUNT(*) AS n FROM artifacts WHERE template_id=?", (key,))["n"]
        if count >= 12:
            fail(422, "每个任务最多附加 12 份材料")
        return public_file(db, add_file(db, user["id"], name, data, ext, "material", template=key))


@router.post("/files")
async def upload_file(user: User, file: Annotated[UploadFile, File()], assignment_id: Annotated[str, Form()] = ""):
    with database() as db:
        writable_task(db, assignment_id, user)
    data = await upload_data(file)
    name, ext = await asyncio.to_thread(verify, file.filename or "", data, SCANNER)
    with database(write=True) as db:
        writable_task(db, assignment_id, user)
        return public_file(db, add_file(db, user["id"], name, data, ext, "uploaded", assignment=assignment_id or None))


class BatchDeleteFiles(BaseModel):
    ids: list[str] = Field(min_length=1, max_length=100)


def _delete_own_file(db, user, key: str):
    row = db.one(
        "SELECT id,owner_id,label,name,extension,parent_id,assignment_id,template_id FROM artifacts WHERE id=?",
        (key,),
    )
    if not row or row["owner_id"] != user["id"] or row["label"] == "material":
        fail(404, "文件不存在或无权删除")
    if db.one("SELECT 1 FROM submission_files WHERE artifact_id=?", (key,)):
        fail(409, "该文件已作为任务成果提交审核，无法删除")
    db.execute("UPDATE artifacts SET parent_id=? WHERE parent_id=?", (row["parent_id"], key))
    db.execute("DELETE FROM office_project_context WHERE artifact_id=?", (key,))
    db.execute("DELETE FROM artifacts WHERE id=? AND owner_id=?", (key, user["id"]))
    audit(db, "file.deleted", user["id"], row["assignment_id"], row["template_id"], artifact_id=key, name=row["name"])
    return row["name"]


@router.post("/files/batch-delete")
def batch_delete_files(body: BatchDeleteFiles, user: User):
    deleted: list[str] = []
    skipped: list[dict[str, str]] = []
    seen: set[str] = set()
    with database(write=True) as db:
        for raw_id in body.ids:
            key = str(raw_id or "").strip()
            if not key or key in seen:
                continue
            seen.add(key)
            row = db.one(
                "SELECT id,owner_id,label,name,parent_id,assignment_id,template_id FROM artifacts WHERE id=?",
                (key,),
            )
            if not row or row["owner_id"] != user["id"] or row["label"] == "material":
                skipped.append({"id": key, "reason": "文件不存在或无权删除"})
                continue
            if db.one("SELECT 1 FROM submission_files WHERE artifact_id=?", (key,)):
                skipped.append({"id": key, "reason": f"「{row['name']}」已提交任务审核，无法删除"})
                continue
            db.execute("UPDATE artifacts SET parent_id=? WHERE parent_id=?", (row["parent_id"], key))
            db.execute("DELETE FROM office_project_context WHERE artifact_id=?", (key,))
            db.execute("DELETE FROM artifacts WHERE id=? AND owner_id=?", (key, user["id"]))
            audit(db, "file.deleted", user["id"], row["assignment_id"], row["template_id"], artifact_id=key, name=row["name"])
            deleted.append(key)
    return {"deleted": deleted, "skipped": skipped}


@router.delete("/files/{key}")
def delete_file(key: str, user: User):
    with database(write=True) as db:
        name = _delete_own_file(db, user, key)
    return {"ok": True, "id": key, "name": name}


@router.get("/files/{key}/download")
def download(key: str, user: User):
    with database(write=True) as db:
        row = file_for(db, user, key)
        audit(db, "file.download", user["id"], row["assignment_id"], row["template_id"], artifact_id=key, sha256=row["digest"])
    name = quote(row["name"], safe="")
    return Response(bytes(row["data"]), media_type=MIMES[row["extension"]], headers={
        "Content-Disposition": "attachment; filename*=UTF-8''" + name,
        "X-Content-Type-Options": "nosniff",
    })


@router.get("/files/{key}/image")
def image(key: str, user: User):
    with database(write=True) as db:
        row = file_for(db, user, key)
        if row["owner_id"] != user["id"]:
            fail(404, "图片不存在")
        if row["extension"] not in {"png", "jpg", "jpeg", "webp"}:
            fail(404, "图片不存在")
        audit(db, "file.image", user["id"], row["assignment_id"], row["template_id"], artifact_id=key)
    return Response(bytes(row["data"]), media_type=MIMES.get(row["extension"], "image/png"), headers={
        "Content-Disposition": "inline",
        "Cache-Control": "private, max-age=0",
    })


@router.get("/files/{key}/preview")
def preview(key: str, user: User):
    with database(write=True) as db:
        row = file_for(db, user, key)
        audit(db, "file.preview", user["id"], row["assignment_id"], row["template_id"], artifact_id=key)
    try:
        text = extract(*file_type(row))
    except Exception:
        fail(422, "文件内容无法提取，请下载后查看")
    return {"id": key, "name": row["name"], "text": text, "parent_id": row["parent_id"], "digest": row["digest"]}


class CopyBody(BaseModel):
    assignment_id: str = Field(min_length=1, max_length=64)


@router.post("/files/{key}/copy")
def copy_to_task(key: str, body: CopyBody, user: User):
    with database(write=True) as db:
        writable_task(db, body.assignment_id, user)
        row = file_for(db, user, key)
        if row["owner_id"] != user["id"] or row["label"] == "material" or row["assignment_id"] and row["assignment_id"] != body.assignment_id:
            fail(403, "只有自己的独立文件可复制到任务")
        if row["assignment_id"] == body.assignment_id:
            return public_file(db, key)
        data, ext = file_type(row)
        return public_file(db, add_file(db, user["id"], row["name"], data, ext, "copied", assignment=body.assignment_id, parent=key))


class EditBody(BaseModel):
    old: str = Field(min_length=1, max_length=300)
    new: str = Field(default="", max_length=500)


@router.post("/files/{key}/edit")
def edit_file(key: str, body: EditBody, user: User):
    with database() as db:
        row = file_for(db, user, key)
        if row["owner_id"] != user["id"] or row["label"] == "material":
            fail(403, "仅可编辑自己添加的文件")
        writable_task(db, row["assignment_id"], user)
    data, count = edit_copy(*file_type(row), body.old, body.new)
    ext = row["extension"]
    stem = row["name"].rsplit(".", 1)[0][:104]
    name, ext = verify(f"{stem} · 修改版.{ext}", data, SCANNER)
    with database(write=True) as db:
        writable_task(db, row["assignment_id"], user)
        edited = add_file(db, user["id"], name, data, ext, "edited", assignment=row["assignment_id"], parent=key)
        return {**public_file(db, edited), "matches": count}


class AgentRequest(BaseModel):
    kind: Literal["docx", "pptx"]
    title: str = Field(min_length=1, max_length=110)
    instructions: str = Field(default="", max_length=6000)
    pages: int = Field(default=8, ge=3, le=20)
    material_ids: list[str] = Field(default_factory=list, max_length=8)


class RenderRequest(BaseModel):
    kind: Literal["docx", "pptx"]
    plan: Plan
    assignment_id: str | None = None
    run_id: str | None = None


class AgentCompletion(BaseModel):
    run_id: str = Field(min_length=1, max_length=100)
    text: str
    usage: dict[str, Any]
    model: str = Field(min_length=1, max_length=200)


class AgentFailure(BaseModel):
    message: str = Field(min_length=1, max_length=1000)


def collect_sources(db, user, task, material_ids):
    data = []
    assignment_id = task["id"] if task else None
    for key in dict.fromkeys(material_ids):
        row = file_for(db, user, key)
        task_material = task and row["label"] == "material" and row["template_id"] == task["template_id"]
        own_upload = row["owner_id"] == user["id"] and row["label"] != "material" and row["assignment_id"] in {assignment_id, None}
        if not (task_material or own_upload):
            fail(403, "请选择当前任务或已上传到工作台的资料")
        if row["extension"] not in {"docx", "pptx", "xlsx", "txt", "md", "csv", "pdf"}:
            fail(422, "当前可提取正文的资料包括 Word、PPT、Excel、PDF、TXT、Markdown 和 CSV")
        data.append((row["name"], *file_type(row)))
    return data


async def read_sources(data):
    sources = []
    text_budget = 20000
    for name, content, ext in data:
        size = min(6500, text_budget)
        text = await asyncio.to_thread(extract, content, ext, size) if size else ""
        sources.append({"name": name, "text": text})
        text_budget -= len(text)
    return sources


async def prepare_agent(user, payload, assignment_id=None, *, engine="desktop"):
    title = clean_text(payload.title, 110)
    with database() as db:
        task = writable_task(db, assignment_id, user)
        data = collect_sources(db, user, task, payload.material_ids)
    sources = await read_sources(data)
    run = uid()
    task_fields = {name: task[name] for name in ("title", "description", "deliverable")} if task else None
    operation, genre = prompting.route(f"{title}\n{payload.instructions}", task_fields)
    prompt = gateway.plan_prompt(payload.kind, title, payload.instructions.strip(), task, sources, payload.pages,
                                 prompting.checklist(operation, genre))
    steps = (["载入任务要求"] if task else [])
    if sources:
        read = sum(bool(source["text"]) for source in sources)
        steps.append(f"读取 {read} 份资料" + (f" · {len(sources)-read} 份没有可提取的正文" if read < len(sources) else ""))
    if label := prompting.steps_label(operation, genre):
        steps.append(label)
    steps.append("正在整理文稿")
    with database(write=True) as db:
        task = writable_task(db, assignment_id, user)
        collect_sources(db, user, task, payload.material_ids)
        db.execute("INSERT INTO runs(id,assignment_id,user_id,kind,engine,status,usage,created_at) VALUES(?,?,?,?,?,?,?,?)",
                   (run, assignment_id, user["id"], payload.kind, engine, "running",
                    pack({"model": None, "prepare": {"pages": payload.pages}}), instant()))
        audit(db, "agent.started", user["id"], assignment_id, run_id=run, engine=engine)
    return {"run_id": run, "system": prompting.plan_system(payload.kind, operation, genre), "prompt": prompt,
            "context": None, "steps": steps, "task": task, "sources": sources,
            "assignment_id": assignment_id, "kind": payload.kind, "pages": payload.pages,
            "engine": engine}


def save_agent_plan(run, user, assignment_id, kind, plan, usage, model, engine):
    if assignment_id:
        with database() as db:
            writable_task(db, assignment_id, user)
    clean_usage = gateway.usage_from(usage)
    recorded = {"model": model, **clean_usage}
    with database(write=True) as db:
        row = db.one("SELECT * FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
        if not row:
            fail(404, "生成任务不存在")
        if row["engine"] != engine:
            fail(404, "生成任务不存在")
        if row["status"] != "running":
            fail(409, "生成任务已结束")
        if row["assignment_id"] != assignment_id or row["kind"] != kind:
            fail(404, "生成任务不存在")
        writable_task(db, assignment_id, user)
        changed = db.execute("UPDATE runs SET status='completed',usage=? WHERE id=? AND status='running'",
                             (pack(recorded), run)).rowcount
        if not changed:
            fail(409, "生成任务已结束")
        audit(db, "agent.completed", user["id"], assignment_id, run_id=run, engine=engine, usage=recorded)
    return {"run_id": run, "plan": plan.model_dump(), "engine": engine, "usage": clean_usage}


def fail_agent_plan(run, user, engine, assignment_id, message):
    with database(write=True) as db:
        changed = db.execute("UPDATE runs SET status='failed' WHERE id=? AND user_id=? AND engine=? AND status='running' AND assignment_id IS ?",
                             (run, user["id"], engine, assignment_id)).rowcount
        if changed:
            audit(db, "agent.failed", user["id"], assignment_id, run_id=run, engine=engine, message=message[:300])
    return bool(changed)


async def agent_work(user, payload, assignment_id=None):
    if gateway.MODE == "desktop":
        fail(409, "请在穹驭办公桌面客户端中使用 AI 办理")
    prepared = await prepare_agent(user, payload, assignment_id, engine=gateway.MODE)
    run = prepared["run_id"]
    try:
        async with PLAN_SLOTS:
            if gateway.MODE == "demo":
                plan, usage = await gateway.plan_for(payload.kind, clean_text(payload.title, 110),
                    payload.instructions.strip(), prepared["task"], prepared["sources"], payload.pages, user["id"], run)
            else:
                text, usage = await gateway.request_model(prepared["prompt"], run,
                    gateway.headers_for(user["id"], assignment_id, run), prepared["system"])
                plan = gateway.response_plan(text, payload.kind, payload.pages)
        return save_agent_plan(run, user, assignment_id, payload.kind, plan, usage,
                               gateway.MODEL if gateway.MODE != "demo" else None, gateway.MODE)
    except BaseException:
        fail_agent_plan(run, user, gateway.MODE, assignment_id, "模型执行失败")
        raise


@router.post("/agent/plan")
async def new_plan(payload: AgentRequest, user: User):
    return await agent_work(user, payload)


@router.post("/assignments/{key}/sessions")
async def task_plan(key: str, payload: AgentRequest, user: User):
    return await agent_work(user, payload, key)


@router.post("/agent/plan/prepare")
async def new_plan_prepare(payload: AgentRequest, user: User):
    if gateway.MODE != "desktop":
        fail(409, "桌面执行模式未启用")
    prepared = await prepare_agent(user, payload)
    return {name: prepared[name] for name in ("run_id", "system", "prompt", "context", "steps")}


@router.post("/assignments/{key}/sessions/prepare")
async def task_plan_prepare(key: str, payload: AgentRequest, user: User):
    if gateway.MODE != "desktop":
        fail(409, "桌面执行模式未启用")
    prepared = await prepare_agent(user, payload, key)
    return {name: prepared[name] for name in ("run_id", "system", "prompt", "context", "steps")}


async def complete_agent_plan(payload: AgentCompletion, user: User, expected_assignment=None, require_unassigned=False):
    with database() as db:
        run = db.one("SELECT * FROM runs WHERE id=? AND user_id=?", (payload.run_id, user["id"]))
        if not run or run["engine"] != "desktop":
            fail(404, "生成任务不存在")
        assignment_id = run["assignment_id"]
        if require_unassigned and assignment_id is not None:
            fail(404, "生成任务不存在")
        if expected_assignment is not None and expected_assignment != assignment_id:
            fail(404, "生成任务不存在")
        if run["status"] != "running":
            fail(409, "生成任务已结束")
        writable_task(db, assignment_id, user)
        options = (unpack(run["usage"]) or {}).get("prepare") or {}
        pages = options.get("pages", 8)
    if len(payload.text.encode("utf-8")) > 256 * 1024:
        fail_agent_plan(payload.run_id, user, "desktop", assignment_id, "模型响应超过 256 KB")
        fail(413, "模型响应超过 256 KB")
    try:
        plan = gateway.response_plan(payload.text, run["kind"], pages)
        return save_agent_plan(payload.run_id, user, assignment_id, run["kind"], plan,
                               payload.usage, payload.model, "desktop")
    except Exception as error:
        fail_agent_plan(payload.run_id, user, "desktop", assignment_id, str(getattr(error, "detail", error)))
        raise


@router.post("/agent/plan/complete")
async def new_plan_complete(payload: AgentCompletion, user: User):
    return await complete_agent_plan(payload, user, require_unassigned=True)


@router.post("/assignments/{key}/sessions/complete")
async def task_plan_complete(key: str, payload: AgentCompletion, user: User):
    return await complete_agent_plan(payload, user, key)


def fail_agent_plan_run(run_id, user, message, expected_assignment=None, require_unassigned=False):
    message = clean_text(message, 1000)
    with database() as db:
        run = db.one("SELECT * FROM runs WHERE id=? AND user_id=?", (run_id, user["id"]))
        if not run or run["engine"] != "desktop":
            fail(404, "生成任务不存在")
        assignment_id = run["assignment_id"]
        if require_unassigned and assignment_id is not None:
            fail(404, "生成任务不存在")
        if expected_assignment is not None and expected_assignment != assignment_id:
            fail(404, "生成任务不存在")
        if run["status"] != "running":
            fail(409, "生成任务已结束")
    if not fail_agent_plan(run_id, user, "desktop", assignment_id, message):
        fail(409, "生成任务已结束")
    return {"ok": True}


@router.post("/agent/plan/{run_id}/fail")
def new_plan_fail(run_id: str, payload: AgentFailure, user: User):
    return fail_agent_plan_run(run_id, user, payload.message, require_unassigned=True)


@router.post("/assignments/{key}/sessions/{run_id}/fail")
def task_plan_fail(key: str, run_id: str, payload: AgentFailure, user: User):
    return fail_agent_plan_run(run_id, user, payload.message, expected_assignment=key)


@router.post("/agent/render")
def render(payload: RenderRequest, user: User):
    def check_run(db):
        writable_task(db, payload.assignment_id, user)
        if payload.run_id:
            run = db.one("SELECT * FROM runs WHERE id=? AND user_id=? AND status='completed'", (payload.run_id, user["id"]))
            if not run or run["assignment_id"] != payload.assignment_id or run["kind"] != payload.kind:
                fail(422, "关联的生成会话无效")
    with database() as db:
        check_run(db)
    plan = payload.plan.clean()
    data = word(plan) if payload.kind == "docx" else powerpoint(plan)
    name = safe_name(plan.title[:105]) + "." + payload.kind
    with database(write=True) as db:
        check_run(db)
        file_id = add_file(db, user["id"], name, data, payload.kind, "generated",
                           assignment=payload.assignment_id, run=payload.run_id)
        return public_file(db, file_id)
