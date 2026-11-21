import asyncio
import base64
import binascii
import datetime
import json
import logging
import re
import struct
from typing import Any, Literal
from uuid import UUID

from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator, model_validator

from . import gateway, prompting, search
from .files import PLAN_SLOTS, collect_sources, public_file, read_sources, writable_task
from .identity import fail
from .office import DeckSpec, Plan, SLIDE_H, SLIDE_W, add_file, clean_text, coerce_deck_spec, powerpoint, render_freeform_pptx, safe_name, word
from .presentation import Replacement, Selection, edit_scope, load_presentation, pptx_file, revise_slide, selected_layout
from .projects import project_for, project_sources
from .store import audit, database, instant, pack, uid, unpack
from .tasks import User, task_for

router = APIRouter(prefix="/api/v1/conversations")
LOG = logging.getLogger("qiongyu.conversations")
MODES = Literal["auto", "reply", "docx", "pptx"]
FORMATS = {"docx": "Word", "pptx": "PowerPoint"}


class NewConversation(BaseModel):
    id: UUID
    assignment_id: str | None = Field(default=None, max_length=100)
    project_id: UUID | None = None


class TurnRequest(BaseModel):
    id: UUID
    message: str = Field(min_length=1, max_length=6000)
    mode: MODES = "auto"
    material_ids: list[str] = Field(default_factory=list, max_length=8)
    selection: Selection | None = None


class Decision(BaseModel):
    model_config = ConfigDict(extra="ignore")
    action: Literal["reply", "export"]
    reply: str = Field(min_length=1, max_length=6500)
    kind: Literal["docx", "pptx"] | None = None
    plan: Plan | None = None
    deck: DeckSpec | None = None


class SlideChange(BaseModel):
    reply: str = Field(min_length=1, max_length=2500)
    changes: list[Replacement] = Field(min_length=1,max_length=45)


class SlidePatch(BaseModel):
    model_config = ConfigDict(extra="ignore")
    reply: str = Field(min_length=1, max_length=6500)
    slide: dict[str, Any]
    design: dict[str, Any] | None = None


class PipelineStage(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    role: Literal["understand", "execute", "accept", "review", "light"]
    model: str = Field(min_length=1, max_length=200)
    outcome: Literal["ok", "format_error", "pass", "fail"]
    usage: dict[str, Any]

    @field_validator("usage", mode="before")
    @classmethod
    def clean_usage(cls, value):
        if not isinstance(value, dict):
            raise ValueError("stage usage must be an object")
        return gateway.usage_from(value)


class PipelineTrace(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    tier: Literal["light", "standard", "critical"]
    verdict: Literal["pass", "fail", "skipped"]
    review: Literal["pass", "fail", "not_needed", "unavailable"]
    issues: list[str] = Field(max_length=10)
    stages: list[PipelineStage] = Field(min_length=1, max_length=8)
    rounds: int = Field(ge=1, le=2)
    version: str = Field(min_length=1, max_length=32)

    @field_validator("issues")
    @classmethod
    def validate_issues(cls, value):
        if any(not isinstance(item, str) or not 1 <= len(item) <= 300 for item in value):
            raise ValueError("pipeline issues must contain 1-300 character strings")
        return value

    @model_validator(mode="after")
    def validate_skipped(self):
        if self.verdict == "skipped" and self.tier != "light":
            raise ValueError("only light pipeline traces may skip acceptance")
        return self


class TurnCompletion(BaseModel):
    text: str
    usage: dict[str, Any]
    model: str = Field(min_length=1, max_length=200)
    pipeline: PipelineTrace | None = None

    model_config = ConfigDict(extra="forbid", strict=True)


class TurnFailure(BaseModel):
    message: str = Field(min_length=1, max_length=1000)
    steps: list[str] | None = Field(default=None, max_length=120)


IMAGE_SIZES = {"1024x1024", "1536x1024", "1024x1536", "1536x864"}
IMAGE_QUALITIES = {"low", "medium", "high"}
PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
MAX_IMAGE_BYTES = 10 * 1024 * 1024
MAX_IMAGE_BASE64 = 4 * ((MAX_IMAGE_BYTES + 2) // 3)
IMAGE_SIZE_DIMENSIONS = {"1024x1024": (1024, 1024), "1536x1024": (1536, 1024),
                         "1024x1536": (1024, 1536), "1536x864": (1536, 864)}
IMAGE_USAGE_FIELDS = ("input_tokens", "output_tokens", "total_tokens")
IMAGE_USAGE_DETAIL_FIELDS = ("image_tokens", "text_tokens")


class ImagePrepare(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    prompt: str = Field(min_length=1, max_length=4000)
    model: str = Field(min_length=1, max_length=200)
    size: Literal["1024x1024", "1536x1024", "1024x1536", "1536x864"]
    quality: Literal["low", "medium", "high"]

    @field_validator("prompt")
    @classmethod
    def trim_prompt(cls, value):
        value = value.strip()
        if not value:
            raise ValueError("prompt must not be empty")
        return value

    @field_validator("model")
    @classmethod
    def validate_image_model(cls, value):
        if not value.startswith("gpt-image-"):
            raise ValueError("model must be an image model")
        return value


class ImageCompletion(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    png_base64: str = Field(min_length=1, max_length=MAX_IMAGE_BASE64)
    width: int = Field(ge=1)
    height: int = Field(ge=1)
    model: str = Field(min_length=1, max_length=200)
    size: Literal["1024x1024", "1536x1024", "1024x1536", "1536x864"]
    quality: Literal["low", "medium", "high"]
    usage: dict[str, Any] = Field(default_factory=dict)

    @field_validator("model")
    @classmethod
    def validate_image_model(cls, value):
        if not value.startswith("gpt-image-"):
            raise ValueError("model must be an image model")
        return value


class ImageFailure(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    message: str = Field(min_length=1, max_length=300)


class TurnImage(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    png_base64: str = Field(min_length=1, max_length=MAX_IMAGE_BASE64)
    width: int = Field(ge=1)
    height: int = Field(ge=1)
    model: str = Field(min_length=1, max_length=200)
    size: Literal["1024x1024", "1536x1024", "1024x1536", "1536x864"]
    quality: Literal["low", "medium", "high"]
    prompt: str = Field(min_length=1, max_length=4000)
    usage: dict[str, Any] = Field(default_factory=dict)

    @field_validator("model")
    @classmethod
    def validate_image_model(cls, value):
        if not value.startswith("gpt-image-"):
            raise ValueError("model must be an image model")
        return value

    @field_validator("prompt")
    @classmethod
    def trim_prompt(cls, value):
        value = value.strip()
        if not value:
            raise ValueError("prompt must not be empty")
        return value


def decode_image(payload):
    try:
        data = base64.b64decode(payload.png_base64, validate=True)
    except (binascii.Error, ValueError):
        fail(422, "生图返回内容无效")
    if (not data or len(data) > MAX_IMAGE_BYTES or base64.b64encode(data).decode("ascii") != payload.png_base64 or
            len(data) < 24 or data[:8] != PNG_SIGNATURE or struct.unpack(">I", data[8:12])[0] != 13 or
            data[12:16] != b"IHDR"):
        fail(422, "生图返回内容无效")
    width, height = struct.unpack(">II", data[16:24])
    expected_width, expected_height = IMAGE_SIZE_DIMENSIONS[payload.size]
    if (width, height) != (payload.width, payload.height) or (width, height) != (expected_width, expected_height):
        fail(422, "生图返回内容无效")
    return data


def image_usage(value):
    if not isinstance(value, dict):
        return {}
    valid = lambda number: isinstance(number, int) and not isinstance(number, bool) and 0 <= number <= 2**53 - 1
    result = {name: value[name] for name in IMAGE_USAGE_FIELDS if valid(value.get(name))}
    for name in ("input_tokens_details", "output_tokens_details"):
        details = value.get(name)
        if isinstance(details, dict):
            cleaned = {field: details[field] for field in IMAGE_USAGE_DETAIL_FIELDS if valid(details.get(field))}
            if cleaned:
                result[name] = cleaned
    return result


def owned_thread(db, key, user):
    row = db.one("SELECT * FROM office_conversations WHERE id=? AND owner_id=?", (key, user["id"]))
    if not row:
        fail(404, "会话不存在")
    if row["assignment_id"]:
        task_for(db, row["assignment_id"], user, "employee")
    if row["project_id"]:
        project_for(db,row["project_id"],user)
    return row


def public_turn(row):
    result = unpack(row["result"])
    body=unpack(row["input"])
    return {"id": row["id"], "message": body["message"], "selection":body.get("selection"), "status": row["status"],
            **{name: result[name] for name in ("reply", "steps", "file", "error", "output_telemetry", "pipeline",
                "kind", "prompt", "model", "size", "quality", "width", "height", "images") if name in result}}


def conversation_summary(row):
    return {name: row[name] for name in ("id", "assignment_id", "project_id", "title", "updated_at")}


@router.get("")
def list_threads(user: User):
    with database() as db:
        rows = db.all("""SELECT c.* FROM office_conversations c LEFT JOIN assignments a ON a.id=c.assignment_id
                       LEFT JOIN office_project_members m ON m.project_id=c.project_id AND m.user_id=c.owner_id
                       WHERE c.owner_id=? AND (c.assignment_id IS NULL OR a.assignee_id=?)
                       AND (c.project_id IS NULL OR m.user_id IS NOT NULL)
                       ORDER BY c.updated_at DESC LIMIT 60""", (user["id"], user["id"]))
    return [conversation_summary(row) for row in rows]


@router.post("")
def new_thread(body: NewConversation, user: User):
    key = str(body.id)
    assignment = body.assignment_id or None
    project=str(body.project_id) if body.project_id else None
    if assignment and project:
        fail(422,"请选择任务会话或项目会话")
    with database(write=True) as db:
        current = db.one("SELECT id FROM office_conversations WHERE id=?", (key,))
        if current:
            row = owned_thread(db, key, user)
            if row["assignment_id"] != assignment or row["project_id"]!=project:
                fail(409, "会话已关联到其他任务或项目")
            return conversation_summary(row)
        task = writable_task(db, assignment, user)
        group=project_for(db,project,user) if project else None
        now = instant()
        db.execute("INSERT INTO office_conversations(id,owner_id,assignment_id,project_id,title,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",
                   (key, user["id"], assignment, project, task["title"] if task else group["name"] if group else "新会话", now, now))
        return conversation_summary(owned_thread(db, key, user))


@router.get("/{key}")
def thread(key: str, user: User, limit: int = Query(50, ge=1, le=100), offset: int = Query(0, ge=0)):
    with database() as db:
        session = owned_thread(db, key, user)
        turns = db.all("SELECT * FROM office_turns WHERE conversation_id=? ORDER BY created_at DESC LIMIT ? OFFSET ?",
                       (key, limit, offset))
    return {**conversation_summary(session), "turns": [public_turn(row) for row in reversed(turns)],
            "has_earlier": len(turns) == limit}


def demo_decision(message, mode, task, sources, recent, latest):
    compact = message.lower()
    powerpoint_request = bool(re.search(r"pptx?|幻灯|演示文稿", compact))
    word_request = bool(re.search(r"\bdocx?\b|word|纪要|通知|报告|方案文档", compact))
    export_request = bool(re.search(r"生成|制作|导出|保存|做一|做个|文件|可编辑|重新", compact))
    if re.search(r"仅回复|只回复|只讨论|先别|暂不|只要大纲", compact):
        export_request = False
    kind = mode if mode in FORMATS else None
    if mode == "auto" and export_request:
        kind = "pptx" if powerpoint_request else "docx" if word_request else (
            latest["kind"] if latest else (task["kind"] if task and task["kind"] in FORMATS else None))
    if kind:
        same_document = latest and latest["kind"] == kind and re.search(r"改|补|调整|重新|上一|同一", compact)
        page_match = re.search(r"(\d{1,2})\s*页", message)
        pages = max(3, min(20, int(page_match[1]))) if page_match else 6
        quoted_title = re.search(r"《([^》]{1,90})》", message)
        title = (latest["plan"]["title"] if same_document else task["title"] if task else
                 quoted_title[1] if quoted_title else "汇报演示文稿" if kind == "pptx" else "工作文稿")
        plan = gateway.demo_plan(kind, title, message, task, sources, pages)
        source = next((item for item in sources if item["text"].strip()), None)
        if source:
            excerpt = f"参考《{source['name'][:18]}》：" + " ".join(source["text"].split())
            limit = 75 if kind == "pptx" else 260
            plan.sections[0].points.append(excerpt[:limit])
        if same_document and not page_match:
            previous = Plan.model_validate(latest["plan"])
            last = previous.sections[-1]
            revision = "本次修改要求：" + " ".join(message.split())[:62 if kind == "pptx" else 200]
            last.points = (last.points[:3] if kind == "pptx" else last.points[:7]) + [revision]
            plan = previous
        note = f"已整理《{plan.title}》的{'修订稿' if same_document else '初稿'}。"
        note += f" 演示文稿包含封面和 {len(plan.sections)} 页正文。" if kind == "pptx" else f" 文稿包含 {len(plan.sections)} 个部分。"
        return Decision(action="export", reply=note, kind=kind, plan=plan)
    lines = []
    if task:
        lines.append(f"当前任务是「{task['title']}」。{task['description']}")
        if task["deliverable"]:
            lines.append("交付要求：" + task["deliverable"])
    for source in sources:
        excerpt = " ".join(source["text"].split())[:320]
        lines.append(f"《{source['name']}》：{excerpt or '未提取到正文。'}")
    if not lines and recent:
        lines.append("上一轮的讨论重点：" + " ".join(recent[-1]["request"].split())[:220])
    if latest:
        headings = "、".join(section["heading"] for section in latest["plan"]["sections"][:7])
        lines.append(f"目前的文稿包含{headings}。")
    if not lines:
        lines.append("请提供要处理的正文、会议笔记或汇报材料。也可以先关联今日任务，我会把主管要求加入会话。")
    return Decision(action="reply", reply="\n\n".join(lines))


def build_turn_request(body, task, sources, recent, latest, project, scope, assignment):
    if scope and scope.get("mode") == "freeform":
        web_search = search.search_status()
        srv_time = web_search.get("server_time") or search.server_time_snapshot()
        slide_copy = dict(scope["slide"])
        indexed_els = []
        for idx, el in enumerate(slide_copy.get("elements") or []):
            if isinstance(el, dict):
                compact = {k: v for k, v in el.items() if v is not None and v is not False}
                compact["_index"] = idx
                if idx in scope["matched_indices"]:
                    compact["_in_selection"] = True
                indexed_els.append(compact)
        slide_copy["elements"] = indexed_els
        sel_info: dict[str, Any] = {
            "页码": scope["page"],
            "总页数": len(scope["deck"].get("slides") or []),
            "选区说明": scope.get("label") or f"第 {scope['page']} 页",
        }
        if scope.get("region"):
            sel_info["圈选区域_英寸"] = scope["region"]
        if scope["matched_indices"]:
            sel_info["命中元素索引"] = scope["matched_indices"]
        if scope.get("hit_bg_image"):
            sel_info["命中全屏背景图"] = True
        if scope.get("selected_objects"):
            sel_info["画布选中的对象摘要"] = scope["selected_objects"]
        payload = {
            "当前服务器时间": {
                "unix_timestamp": srv_time["unix_timestamp"],
                "date": srv_time["date"],
                "year": srv_time["year"],
                "display": srv_time["display"],
            },
            "员工修改要求": body.message,
            "选区范围": sel_info,
            "整套PPT设计系统": scope["deck"].get("design") or {},
            "当前页完整画布": slide_copy,
            "已有图片资产": scope["base_image_meta"],
            "近期会话": recent[-4:],
            "参考资料": sources,
        }
        plugin_context = {"项目": project, "资料": sources} if project else None
        if project and gateway.MODE in {"dsh", "desktop"}:
            payload["项目"] = project
            payload["参考资料"] = []
            payload["接入说明"] = "如需查阅项目资料，可先调用 office_context_read。"
        return {
            "kind": "slide_freeform",
            "system": prompting.SLIDE_FREEFORM_SYSTEM,
            "prompt": pack(payload),
            "context": plugin_context,
            "preset_step": None,
            "wants_pptx": True,
            "image_index_offset": len(scope["base_image_ids"]),
            "web_search": {
                "enabled": bool(web_search["configured"]),
                "required": False,
                "server_time": srv_time,
            },
        }
    if scope:
        context = {"选中页": scope["page"], "可编辑文本框": scope["boxes"], "项目": project, "参考资料": sources}
        prompt = {"请求": body.message, "选区资料": context}
        if gateway.MODE in {"dsh", "desktop"}:
            prompt = {"请求": body.message, "说明": "先调用 office_context_read 获取选区与资料。"}
        return {"kind": "slide", "system": prompting.SLIDE_SYSTEM, "prompt": pack(prompt), "context": context,
                "preset_step": None, "wants_pptx": False, "image_index_offset": 0,
                "web_search": {"enabled": False, "required": False}}

    task_fields = {name: task[name] for name in ("title", "description", "deliverable", "due_at")} if task else None
    plan = latest.get("plan") if latest else None
    latest_kind = latest.get("kind") if latest else None
    include_pptx = prompting.wants_pptx(body.message, body.mode, task, latest_kind)
    operation, genre = prompting.route(body.message, task_fields,
                                       plan.get("title", "") if isinstance(plan, dict) else "",
                                       [item["request"] for item in recent])
    web_search = search.search_status()
    srv_time = web_search.get("server_time") or search.server_time_snapshot()
    require_search = bool(web_search["configured"] and web_search["required"])
    payload = {
        "当前服务器时间": {
            "unix_timestamp": srv_time["unix_timestamp"],
            "iso": srv_time["iso"],
            "date": srv_time["date"],
            "year": srv_time["year"],
            "month": srv_time["month"],
            "display": srv_time["display"],
            "搜索与叙事基准": (
                f"当前服务器真实时间为 {srv_time['year']} 年 {srv_time['month']} 月（今日 {srv_time['date']}，Unix时间戳 {srv_time['unix_timestamp']}）。"
                f"所有“最新/今年/目前/近期”均指 {srv_time['year']} 年，调用 web_search 检索时必须优先搜索 {srv_time['year']} 年最新进展，严禁将 2024 或 2025 年当作当前年份！"
            ),
        },
        "当前请求": body.message,
        "近期会话": recent,
        "当前任务": task_fields,
        "参考资料": sources,
        "最近文稿": {key: latest[key] for key in ("plan", "kind")} if latest else None,
        "输出偏好": body.mode,
    }
    rules = prompting.checklist(operation, genre)
    if rules:
        payload["写作规范"] = rules
    if pages := prompting.page_rule(body.message, body.mode, latest_kind):
        payload["页数要求"] = pages
    plugin_context = {"项目": project, "资料": sources} if project else None
    if project and gateway.MODE in {"dsh", "desktop"}:
        payload["项目"] = project
        payload["参考资料"] = []
        payload["接入说明"] = "先调用 office_context_read，获取本轮授权资料。"
    return {"kind": "conversation",
            "system": prompting.conversation_system(body.mode, operation, genre,
                                                    include_pptx=include_pptx,
                                                    require_web_search=require_search,
                                                    server_time=srv_time),
            "prompt": pack(payload), "context": plugin_context,
            "preset_step": prompting.steps_label(operation, genre),
            "wants_pptx": include_pptx,
            "web_search": {
                "enabled": bool(web_search["configured"]),
                "required": require_search,
                "server_time": srv_time,
            }}


def _repair_json_text(text: str) -> str:
    out: list[str] = []
    in_string = False
    escaped = False
    stack: list[str] = []
    for ch in text:
        code = ord(ch)
        if in_string:
            if escaped:
                out.append(ch)
                escaped = False
            elif ch == "\\":
                out.append(ch)
                escaped = True
            elif ch == '"':
                out.append(ch)
                in_string = False
            elif ch == "\n":
                out.append("\\n")
            elif ch == "\r":
                out.append("\\r")
            elif ch == "\t":
                out.append("\\t")
            elif code < 0x20:
                out.append(" ")
            else:
                out.append(ch)
        else:
            if ch == '"':
                in_string = True
                out.append(ch)
            elif ch == "{":
                stack.append("}")
                out.append(ch)
            elif ch == "[":
                stack.append("]")
                out.append(ch)
            elif ch in ("}", "]") and stack and stack[-1] == ch:
                stack.pop()
                out.append(ch)
            else:
                out.append(ch)
    if in_string:
        out.append('"')
    joined = re.sub(r",\s*([}\]])", r"\1", "".join(out))
    joined = re.sub(r",\s*$", "", joined)
    if stack:
        joined += "".join(reversed(stack))
    return re.sub(r",\s*([}\]])", r"\1", joined)


def _extract_json_object(text: str) -> dict[str, Any]:
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", text.strip(), flags=re.I)
    candidates = [cleaned]
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start != -1 and end > start:
        candidates.append(cleaned[start:end + 1])
    if start != -1:
        candidates.append(cleaned[start:])
    for cand in candidates:
        for variant in (cand, re.sub(r",\s*([}\]])", r"\1", cand), _repair_json_text(cand)):
            try:
                parsed = json.loads(variant)
                if isinstance(parsed, dict):
                    return parsed
            except json.JSONDecodeError:
                pass
    plain = re.sub(r"^\s*```(?:markdown|md|text)?\s*|\s*```\s*$", "", text.strip(), flags=re.I).strip()
    if len(plain) >= 2 and not plain.startswith("{") and '"slides"' not in plain:
        return {"action": "reply", "reply": plain[:6500]}
    fail(502, "办公助理返回了无法读取的结果，请重试当前请求")


def parse_turn_response(text, kind):
    cleaned = re.sub(r"^\s*```(?:json)?\s*|\s*```\s*$", "", text.strip(), flags=re.I)
    if kind == "slide":
        try:
            return SlideChange.model_validate_json(cleaned)
        except ValidationError as exc:
            LOG.warning("SlideChange validation failed: %s", exc)
            fail(502, "模型未返回可执行的文字修改，请重试")
    data = _extract_json_object(text)
    if kind == "slide_freeform":
        raw_slide = data.get("slide")
        if isinstance(raw_slide, dict):
            reply = str(data.get("reply") or "").strip()[:6500] or "已按要求修改选中页面。"
            design = data.get("design") if isinstance(data.get("design"), dict) else None
            return SlidePatch(reply=reply, slide=raw_slide, design=design)
        # 兼容模型在选区修改时输出整套 deck 的情况
        raw_deck = data.get("deck")
        if isinstance(raw_deck, dict) and isinstance(raw_deck.get("slides"), list) and raw_deck["slides"]:
            deck = coerce_deck_spec(raw_deck)
            reply = str(data.get("reply") or "").strip()[:6500] or f"已按要求更新演示文稿《{deck.title}》。"
            return Decision(action="export", reply=reply, kind="pptx", deck=deck)
        fail(502, "模型未返回有效的单页修改方案，请重试")
    raw_kind = str(data.get("kind") or "").strip().lower()
    if raw_kind in {"ppt", "powerpoint"}:
        data["kind"] = "pptx"
    elif raw_kind in {"doc", "word"}:
        data["kind"] = "docx"
    raw_deck = data.get("deck")
    if not isinstance(raw_deck, dict) and isinstance(data.get("plan"), dict) and isinstance(data["plan"].get("slides"), list):
        raw_deck = data["plan"]
    if not isinstance(raw_deck, dict) and isinstance(data.get("slides"), list):
        raw_deck = data
    if (data.get("action") == "export" or data.get("kind") == "pptx") and isinstance(raw_deck, dict) and isinstance(raw_deck.get("slides"), list):
        deck = coerce_deck_spec(raw_deck)
        reply = str(data.get("reply") or "").strip()[:6500] or f"已生成演示文稿《{deck.title}》。"
        return Decision(action="export", reply=reply, kind="pptx", deck=deck)
    try:
        return Decision.model_validate(data)
    except ValidationError as exc:
        LOG.warning("Decision validation failed: %s", exc)
        fail(502, "办公助理返回了无法读取的结果，请重试当前请求")


async def decide(body, task, sources, recent, latest, user, run, project=None):
    if gateway.MODE == "demo":
        return demo_decision(body.message, body.mode, task, sources, recent, latest), {}
    prepared = build_turn_request(body, task, sources, recent, latest, project, None,
                                  task["id"] if task else None)
    text, usage = await gateway.request_model(prepared["prompt"], run,
        gateway.headers_for(user["id"], task["id"] if task else None, run), prepared["system"], prepared["context"])
    return parse_turn_response(text, prepared["kind"]), usage


async def decide_slide(body, scope, sources, project, assignment, user, run):
    if gateway.MODE == "demo":
        box = next((b for b in scope["boxes"] if b["text"] != "穹驭办公" and not b["text"].isdigit()), scope["boxes"][0])
        match = re.search(r"[「“《]([^」”》]+)[」”》]", body.message)
        text = match[1] if match else body.message.partition("改为")[2].strip() or body.message[:80]
        return SlideChange(reply=f"已修改第 {scope['page']} 页选中范围内的文字。",
                           changes=[Replacement(shape_id=box["shape_id"], text=text)]), {}
    prepared = build_turn_request(body, None, sources, [], None, project, scope, assignment)
    raw, usage = await gateway.request_model(prepared["prompt"], run,
        gateway.headers_for(user["id"], assignment, run), prepared["system"], prepared["context"])
    return parse_turn_response(raw, prepared["kind"]), usage


def office_file(decision, mode, image_blobs=()):
    if decision.action == "reply":
        return None
    if not decision.kind or mode == "reply" or (mode in FORMATS and mode != decision.kind):
        fail(502, "返回的文稿格式与本轮请求不一致")
    kind = decision.kind
    if kind == "pptx" and decision.deck is not None:
        deck = decision.deck.clean()
        name = safe_name(deck.title[:105]) + ".pptx"
        data = render_freeform_pptx(deck, image_blobs)
        summary = deck.model_dump()
        return kind, summary, name, data
    if not decision.plan:
        fail(502, "返回的文稿缺少内容结构")
    plan = decision.plan.clean()
    # 兼容旧 plan 格式兜底：仅拒绝单条长到任何字号都放不下的情况。
    if kind == "pptx" and any(len(" ".join(p.split())) > 200 for s in plan.sections for p in s.points):
        fail(502, "有要点超过 200 字，超出 PPT 版式，请缩短后重试")
    name = safe_name(plan.title[:105]) + "." + kind
    data = word(plan) if kind == "docx" else powerpoint(plan)
    return kind, plan.model_dump(), name, data


def event(kind, **payload):
    return pack({"type": kind, **payload}) + "\n"


def stream(generator):
    return StreamingResponse(generator, media_type="application/x-ndjson", headers={"X-Accel-Buffering": "no"})


async def replay(row):
    yield event("complete", turn=public_turn(row))


def fail_turn(run, user, assignment, message, steps=None):
    with database(write=True) as db:
        current = db.one("SELECT result FROM office_turns WHERE id=? AND status='running'", (run,))
        if not current:
            return
        result = unpack(current["result"] or "{}")
        result["error"] = message
        if isinstance(steps, list) and steps:
            existing = list(result.get("steps") or [])
            for step in steps:
                if not isinstance(step, str):
                    continue
                cleaned_step = clean_text(step, 240)
                if cleaned_step and (not existing or existing[-1] != cleaned_step):
                    existing.append(cleaned_step)
            result["steps"] = existing[-80:]
        changed = db.execute("UPDATE office_turns SET status='failed',result=? WHERE id=? AND status='running'",
                             (pack(result), run)).rowcount
        if changed:
            db.execute("UPDATE runs SET status='failed' WHERE id=?", (run,))
            run_row = db.one("SELECT engine FROM runs WHERE id=?", (run,))
            audit(db, "agent.failed", user["id"], assignment, run_id=run,
                  engine=run_row["engine"] if run_row else gateway.MODE)
        return bool(changed)


def history_for(db, key):
    history = db.all("""SELECT input,result FROM office_turns WHERE conversation_id=? AND status='completed'
                       ORDER BY created_at DESC LIMIT 12""", (key,))
    recent, latest = [], None
    for index, row in enumerate(history):
        result = unpack(row["result"])
        if index < 8 and isinstance(result.get("reply"), str):
            recent.append({"request": unpack(row["input"])["message"][:1500], "reply": result["reply"][:1500]})
        file = result.get("file")
        if (result.get("plan") and isinstance(file, dict) and file.get("extension") in {"docx", "pptx"}
                and latest is None):
            latest = {name: result[name] for name in ("plan", "kind", "file")}
    recent.reverse()
    return recent, latest


def _boxes_intersect(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> bool:
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return ax < bx + bw and ax + aw > bx and ay < by + bh and ay + ah > by


def _collect_turn_image_assets(db, key: str, source_run_id: str | None, source_result: dict[str, Any], user_id: str):
    """提取某份 PPT 依赖的有序图片 artifact IDs 与可读描述，确保多轮圈选换图时 image_index 1..N 精确对齐。"""
    image_meta_by_id: dict[str, dict[str, Any]] = {}
    turns = db.all(
        "SELECT id, result FROM office_turns WHERE conversation_id=? AND status IN ('completed','failed') ORDER BY created_at DESC LIMIT 20",
        (key,),
    )
    for row in turns:
        res = unpack(row["result"] or "{}")
        for img in res.get("images") or []:
            fid = (img.get("file") or {}).get("id") if isinstance(img, dict) else None
            if fid and fid not in image_meta_by_id:
                image_meta_by_id[fid] = {
                    "size": img.get("size") or "",
                    "prompt": str(img.get("prompt") or "")[:180],
                }

    ordered_ids: list[str] = []
    if isinstance(source_result.get("deck_image_ids"), list) and source_result["deck_image_ids"]:
        ordered_ids = [str(fid) for fid in source_result["deck_image_ids"] if isinstance(fid, str) and fid]
    elif isinstance(source_result.get("images"), list) and source_result["images"]:
        for img in source_result["images"]:
            fid = (img.get("file") or {}).get("id") if isinstance(img, dict) else None
            if fid:
                ordered_ids.append(fid)
    else:
        for row in turns:
            res = unpack(row["result"] or "{}")
            imgs = res.get("images") or []
            if imgs:
                for img in imgs:
                    fid = (img.get("file") or {}).get("id") if isinstance(img, dict) else None
                    if fid:
                        ordered_ids.append(fid)
                break

    valid_ids: list[str] = []
    meta_list: list[dict[str, Any]] = []
    for fid in ordered_ids:
        art = db.one("SELECT id FROM artifacts WHERE id=? AND owner_id=?", (fid, user_id))
        if not art:
            continue
        valid_ids.append(fid)
        info = image_meta_by_id.get(fid, {})
        meta_list.append({
            "image_index": len(valid_ids),
            "size": info.get("size") or "1536x864",
            "prompt": info.get("prompt") or "已有配图",
        })
    return valid_ids, meta_list


def _load_image_blobs_by_ids(db, file_ids: list[str], user_id: str) -> list[bytes]:
    blobs: list[bytes] = []
    for fid in file_ids:
        art = db.one("SELECT data FROM artifacts WHERE id=? AND owner_id=?", (fid, user_id))
        if art and art["data"]:
            blobs.append(bytes(art["data"]))
    return blobs


def _is_freeform_deck(plan: Any) -> bool:
    if not isinstance(plan, dict) or not isinstance(plan.get("slides"), list) or not plan["slides"]:
        return False
    return any(
        isinstance(s, dict) and (isinstance(s.get("elements"), list) and len(s["elements"]) > 0 or s.get("bg_image") is not None)
        for s in plan["slides"]
    )


def selection_target(db, key, body, user):
    if not body.selection:
        return None
    row = pptx_file(db, body.selection.file_id, user)
    turn_row = db.one("SELECT id, result FROM office_turns WHERE id=? AND conversation_id=? AND status='completed'",
                      (row["run_id"], key))
    if not turn_row:
        fail(403, "只能修改当前会话生成的 PPT")
    source_result = unpack(turn_row["result"] or "{}")
    plan = source_result.get("plan")
    freeform = _is_freeform_deck(plan) and 1 <= body.selection.slide <= len(plan["slides"])
    base_image_ids: list[str] = []
    base_image_meta: list[dict[str, Any]] = []
    if freeform:
        base_image_ids, base_image_meta = _collect_turn_image_assets(db, key, turn_row["id"], source_result, user["id"])
    return {
        "id": row["id"],
        "name": row["name"],
        "data": bytes(row["data"]),
        "freeform": freeform,
        "deck": plan if freeform else None,
        "base_image_ids": base_image_ids,
        "base_image_meta": base_image_meta,
    }


def build_freeform_scope(target: dict[str, Any], selection: Selection) -> dict[str, Any]:
    deck = target["deck"]
    slide_index = selection.slide - 1
    slide_dict = deck["slides"][slide_index]
    elements = slide_dict.get("elements") if isinstance(slide_dict.get("elements"), list) else []

    region_dict = None
    query_boxes: list[tuple[float, float, float, float]] = []
    selected_objects: list[dict[str, Any]] = []
    hit_bg_image = False

    if selection.region is not None:
        r = selection.region
        region_dict = {"x": round(r.x, 2), "y": round(r.y, 2), "w": round(r.w, 2), "h": round(r.h, 2)}
        query_boxes.append((r.x, r.y, r.w, r.h))
        if slide_dict.get("bg_image") is not None:
            hit_bg_image = True

    picked_ids = list(dict.fromkeys(
        ([selection.shape_id] if selection.shape_id else []) + list(selection.shape_ids or [])
    ))
    if picked_ids:
        try:
            layout = selected_layout(load_presentation(target["data"]), selection.slide)
            scale_to_in = SLIDE_W / 1000.0
            for obj in layout.get("objects") or []:
                if obj.get("id") not in picked_ids:
                    continue
                bx = float(obj.get("left", 0.0)) * scale_to_in
                by = float(obj.get("top", 0.0)) * scale_to_in
                bw = max(0.05, float(obj.get("width", 0.0)) * scale_to_in)
                bh = max(0.05, float(obj.get("height", 0.0)) * scale_to_in)
                if not query_boxes or selection.region is None:
                    query_boxes.append((bx, by, bw, bh))
                if obj.get("kind") == "image" and obj.get("is_background") and slide_dict.get("bg_image") is not None:
                    hit_bg_image = True
                selected_objects.append({
                    "kind": obj.get("kind") or "shape",
                    "x": round(bx, 2), "y": round(by, 2), "w": round(bw, 2), "h": round(bh, 2),
                    **({"text": str(obj["text"])[:120]} if obj.get("text") else {}),
                    **({"is_background": True} if obj.get("is_background") else {}),
                })
        except Exception:
            pass

    matched_indices: list[int] = []
    if query_boxes:
        for idx, el in enumerate(elements):
            if not isinstance(el, dict):
                continue
            ex = float(el.get("x") or 0.0)
            ey = float(el.get("y") or 0.0)
            ew = max(0.05, float(el.get("w") or 0.2))
            eh = max(0.05, float(el.get("h") or 0.2))
            if any(_boxes_intersect((ex, ey, ew, eh), qb) for qb in query_boxes):
                matched_indices.append(idx)

    return {
        "mode": "freeform",
        "page": selection.slide,
        "slide_index": slide_index,
        "deck": deck,
        "slide": slide_dict,
        "region": region_dict,
        "matched_indices": matched_indices,
        "hit_bg_image": hit_bg_image,
        "selected_objects": selected_objects[:20],
        "label": selection.label or "",
        "base_image_ids": target["base_image_ids"],
        "base_image_meta": target["base_image_meta"],
    }


def apply_freeform_slide_patch(target: dict[str, Any], scope: dict[str, Any], decision: SlidePatch | Decision, image_blobs: list[bytes]):
    if isinstance(decision, Decision) and decision.deck is not None:
        updated_deck = decision.deck.clean()
    elif isinstance(decision, SlidePatch):
        base_deck_raw = dict(target["deck"])
        merged_design = dict(base_deck_raw.get("design") or {})
        if isinstance(decision.design, dict):
            if isinstance(decision.design.get("palette"), dict):
                base_pal = dict(merged_design.get("palette") or {})
                base_pal.update(decision.design["palette"])
                merged_design["palette"] = base_pal
            if decision.design.get("concept"):
                merged_design["concept"] = decision.design["concept"]
            if decision.design.get("font"):
                merged_design["font"] = decision.design["font"]
        # 仅把新单页通过 coerce_deck_spec 归一化（支持自动百分比/像素坐标系纠偏），其他页原样保留
        coerced_single = coerce_deck_spec({
            "title": base_deck_raw.get("title") or "演示文稿",
            "subtitle": base_deck_raw.get("subtitle") or "",
            "design": merged_design,
            "slides": [decision.slide],
        })
        new_slides = list(base_deck_raw.get("slides") or [])
        new_slides[scope["slide_index"]] = coerced_single.slides[0].model_dump()
        updated_deck = coerce_deck_spec({
            "title": base_deck_raw.get("title") or "演示文稿",
            "subtitle": base_deck_raw.get("subtitle") or "",
            "design": coerced_single.design.model_dump(),
            "slides": new_slides,
        }).clean()
    else:
        fail(502, "模型未返回有效的单页修改方案，请重试")

    name = target["name"] or (safe_name(updated_deck.title[:105]) + ".pptx")
    data = render_freeform_pptx(updated_deck, image_blobs)
    return "pptx", updated_deck.model_dump(), name, data


def stale_before():
    # 桌面端智能分工全程上限 30 分钟（desktop/lib/pipeline.cjs DEADLINE_MS），这里必须更长，
    # 否则仍在执行的一轮会被下一次请求判成"已中断"。
    return instant(datetime.datetime.now(datetime.UTC) - datetime.timedelta(minutes=40))


def check_existing(db, key, body, request, allow_replay, stale):
    existing = db.one("SELECT * FROM office_turns WHERE conversation_id=? AND request_id=?", (key, str(body.id)))
    if existing:
        if existing["input"] != request:
            fail(409, "本轮请求已经提交，请等待执行完成")
        if existing["status"] == "running" and existing["created_at"] >= stale:
            fail(409, "本轮请求已经提交，请等待执行完成")
        if allow_replay and existing["status"] != "running":
            return existing
        fail(409, "本轮请求已经提交，请等待执行完成")
    if db.one("SELECT id FROM office_turns WHERE conversation_id=? AND status='running' AND created_at>=?",
              (key, stale)):
        fail(409, "当前会话有一轮任务正在执行")
    return None


async def prepare_conversation(key, body, user, *, engine="desktop", allow_replay=False):
    body.message = clean_text(body.message, 6000)
    request = pack(body.model_dump(mode="json", exclude={"selection"} if body.selection is None else set()))
    stale = stale_before()
    with database() as db:
        session = owned_thread(db, key, user)
        assignment = session["assignment_id"]
        replay_row = check_existing(db, key, body, request, allow_replay, stale)
        if replay_row:
            return {"replay": replay_row}
        task = writable_task(db, assignment, user)
        data = collect_sources(db, user, task, body.material_ids)
        project_data = project_sources(db, session["project_id"], user)
        target = selection_target(db, key, body, user)
        recent, latest = history_for(db, key)

    project, notes, shared_files = project_data
    sources = await read_sources(data)
    if project:
        sources.extend([*notes, *await read_sources(shared_files)])
    if target:
        scope = await asyncio.to_thread(
            build_freeform_scope if target.get("freeform") else edit_scope,
            target if target.get("freeform") else target["data"],
            body.selection,
        )
    else:
        scope = None
    prompt_data = build_turn_request(body, task, sources, recent, latest, project, scope, assignment)
    steps = []
    if task:
        steps.append("载入任务要求")
    if sources:
        read = sum(bool(source["text"]) for source in sources)
        steps.append(f"读取 {read} 份资料" + (f" · {len(sources)-read} 份没有可提取的正文" if read < len(sources) else ""))
    if project:
        steps.append(f"载入项目「{project}」的共享资料")
    if scope:
        if scope.get("mode") == "freeform":
            desc = body.selection.label or ("圈选区域" if body.selection.region else ("选中元素" if (body.selection.shape_id or body.selection.shape_ids) else "整页自由画布"))
            steps.append(f"已锁定第 {scope['page']} 页 · {desc}")
        else:
            steps.append(f"已锁定第 {scope['page']} 页" + ("选中的文本框" if body.selection.shape_id else "的可编辑文字"))
    if prompt_data["preset_step"]:
        steps.append(prompt_data["preset_step"])

    run, now = uid(), instant()
    with database(write=True) as db:
        session = owned_thread(db, key, user)
        if session["assignment_id"] != assignment:
            fail(409, "会话关联已变化，请重新打开会话")
        replay_row = check_existing(db, key, body, request, allow_replay, stale)
        if replay_row:
            return {"replay": replay_row}
        for prior in db.all("SELECT id,result FROM office_turns WHERE conversation_id=? AND status='running' AND created_at<?",
                            (key, stale)):
            result = unpack(prior["result"] or "{}")
            result["error"] = "上次执行已中断，请重新发送"
            db.execute("UPDATE office_turns SET status='failed',result=? WHERE id=? AND status='running'",
                       (pack(result), prior["id"]))
            db.execute("UPDATE runs SET status='failed' WHERE id=?", (prior["id"],))
        if db.one("SELECT id FROM office_turns WHERE conversation_id=? AND status='running'", (key,)):
            fail(409, "当前会话有一轮任务正在执行")
        task = writable_task(db, assignment, user)
        project_sources(db, session["project_id"], user)
        if body.selection:
            selection_target(db, key, body, user)
        recent, latest = history_for(db, key)
        prompt_data = build_turn_request(body, task, sources, recent, latest, project, scope, assignment)
        db.execute("INSERT INTO runs(id,assignment_id,user_id,kind,engine,status,usage,created_at) VALUES(?,?,?,?,?,?,?,?)",
                   (run, assignment, user["id"], "conversation", engine, "running", pack({"model": None}), now))
        db.execute("INSERT INTO office_turns(id,conversation_id,request_id,input,result,status,created_at) VALUES(?,?,?,?,?,?,?)",
                   (run, key, str(body.id), request, pack({"steps": steps}), "running", now))
        title = " ".join(body.message.split())[:42] if session["title"] == "新会话" else session["title"]
        db.execute("UPDATE office_conversations SET title=?,updated_at=? WHERE id=?", (title, now, key))
        audit(db, "agent.started", user["id"], assignment, run_id=run, engine=engine)
    return {"run_id": run, "kind": prompt_data["kind"], "system": prompt_data["system"],
            "prompt": prompt_data["prompt"], "context": prompt_data["context"],
            "steps": [*steps, "正在整理回复"], "wants_pptx": bool(prompt_data.get("wants_pptx")),
            "image_index_offset": int(prompt_data.get("image_index_offset") or 0),
            "web_search": prompt_data.get("web_search") or {"enabled": False, "required": False},
            "body": body, "user": user, "assignment": assignment,
            "task": task, "sources": sources, "recent": recent, "latest": latest,
            "project": project, "scope": scope, "target": target, "engine": engine,
            "conversation_id": key}


async def complete_turn(key, run, payload, user, *, expected_engine):
    image_blobs = []
    deck_image_ids: list[str] = []
    with database() as db:
        turn = db.one("SELECT * FROM office_turns WHERE id=? AND conversation_id=?", (run, key))
        run_row = db.one("SELECT * FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
        if not turn or not run_row:
            fail(404, "本轮执行不存在")
        if run_row["engine"] != expected_engine:
            fail(404, "本轮执行不存在")
        if turn["status"] != "running" or run_row["status"] != "running":
            fail(409, "本轮执行已结束")
        session = owned_thread(db, key, user)
        assignment = session["assignment_id"]
        task = writable_task(db, assignment, user)
        body = TurnRequest.model_validate(unpack(turn["input"]))
        target = selection_target(db, key, body, user)
        recent, latest = history_for(db, key)
        turn_result = unpack(turn["result"] or "{}")
        current_run_image_ids: list[str] = []
        for img_entry in turn_result.get("images") or []:
            fid = (img_entry.get("file") or {}).get("id") if isinstance(img_entry, dict) else None
            if fid:
                current_run_image_ids.append(fid)

        if target and target.get("freeform"):
            # 自由画布局部修改：严格拼接 原PPT有序历史图 + 本轮新生成图，保证 1..M 原图不变、M+1.. 新图精确对齐
            deck_image_ids = [*target.get("base_image_ids", []), *current_run_image_ids]
            image_blobs = _load_image_blobs_by_ids(db, deck_image_ids, user["id"])
        elif current_run_image_ids:
            deck_image_ids = current_run_image_ids
            image_blobs = _load_image_blobs_by_ids(db, deck_image_ids, user["id"])
        else:
            prior_turns = db.all(
                "SELECT result FROM office_turns WHERE conversation_id=? AND id!=? AND status IN ('completed','failed') ORDER BY created_at DESC LIMIT 8",
                (key, run),
            )
            for p_row in prior_turns:
                p_res = unpack(p_row["result"] or "{}")
                if isinstance(p_res.get("deck_image_ids"), list) and p_res["deck_image_ids"]:
                    deck_image_ids = [str(fid) for fid in p_res["deck_image_ids"] if isinstance(fid, str) and fid]
                else:
                    for img_entry in p_res.get("images") or []:
                        fid = (img_entry.get("file") or {}).get("id") if isinstance(img_entry, dict) else None
                        if fid:
                            deck_image_ids.append(fid)
                if deck_image_ids:
                    image_blobs = _load_image_blobs_by_ids(db, deck_image_ids, user["id"])
                    if image_blobs:
                        break
                    deck_image_ids = []

    if len(payload.text.encode("utf-8")) > 512 * 1024:
        fail_turn(run, user, assignment, "模型响应超过 512 KB")
        fail(413, "模型响应超过 512 KB")
    try:
        if target and target.get("freeform"):
            scope = await asyncio.to_thread(build_freeform_scope, target, body.selection)
            decision = parse_turn_response(payload.text, "slide_freeform")
            revised = None
            rendered = await asyncio.to_thread(apply_freeform_slide_patch, target, scope, decision, image_blobs)
        elif target:
            scope = await asyncio.to_thread(edit_scope, target["data"], body.selection)
            decision = parse_turn_response(payload.text, "slide")
            revised = await asyncio.to_thread(revise_slide, target["data"], scope, decision.changes)
            rendered = None
        else:
            scope = None
            decision = parse_turn_response(payload.text, "conversation")
            revised = None
            rendered = await asyncio.to_thread(office_file, decision, body.mode, image_blobs)
    except HTTPException as error:
        fail_turn(run, user, assignment, str(error.detail))
        raise
    except Exception:
        LOG.exception("Office turn response rejected on run %s", run)
        message = "模型返回的内容无法处理，请重试当前请求"
        fail_turn(run, user, assignment, message)
        fail(502, message)

    stored_usage = {"model": payload.model, **gateway.usage_from(payload.usage)}
    output_telemetry = {key: stored_usage[key] for key in
                        ("max_tokens_clamped", "finish_reasons", "truncated") if key in stored_usage}
    with database(write=True) as db:
        owned_thread(db, key, user)
        writable_task(db, assignment, user)
        current = db.one("SELECT status,result FROM office_turns WHERE id=? AND conversation_id=?", (run, key))
        current_run = db.one("SELECT status,engine FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
        if not current or not current_run or current_run["engine"] != expected_engine:
            fail(404, "本轮执行不存在")
        if current["status"] != "running" or current_run["status"] != "running":
            fail(409, "本轮执行已结束")
        result = unpack(current["result"] or "{}")
        steps = result.get("steps", [])
        steps.append("本机办公客户端返回办公方案" if expected_engine == "desktop" else
                     "运行本地办公演示" if expected_engine == "demo" else
                     "DeepSeek Harness 返回办公方案" if expected_engine == "dsh" else "公司模型网关返回办公方案")
        result.update(reply=clean_text(decision.reply, 6500), steps=steps)
        if output_telemetry:
            result["output_telemetry"] = output_telemetry
        if payload.pipeline is not None:
            result["pipeline"] = payload.pipeline.model_dump()
        if revised:
            artifact = add_file(db, user["id"], target["name"], revised, "pptx", "generated",
                                assignment=assignment, parent=target["id"], run=run)
            result.update(file=public_file(db, artifact), kind="pptx")
            steps.append(f"已保存第 {scope['page']} 页的修订版")
        elif rendered:
            kind, plan_data, name, content = rendered
            parent = target["id"] if (target and target.get("freeform")) else (latest["file"]["id"] if latest and latest["file"]["name"] == name else None)
            artifact = add_file(db, user["id"], name, content, kind, "generated",
                                assignment=assignment, parent=parent, run=run)
            result.update(file=public_file(db, artifact), plan=plan_data, kind=kind)
            if kind == "pptx" and deck_image_ids:
                result["deck_image_ids"] = deck_image_ids
            steps.append(f"已保存第 {scope['page']} 页的精修版" if (scope and scope.get("mode") == "freeform") else ("已导出 " + FORMATS[kind]))
        changed = db.execute("UPDATE office_turns SET result=?,status='completed' WHERE id=? AND status='running'",
                             (pack(result), run)).rowcount
        if not changed:
            fail(409, "本轮执行已结束")
        db.execute("UPDATE runs SET status='completed',usage=? WHERE id=? AND status='running'",
                   (pack(stored_usage), run))
        db.execute("UPDATE office_conversations SET updated_at=? WHERE id=?", (instant(), key))
        audit(db, "agent.completed", user["id"], assignment, run_id=run, engine=expected_engine, usage=stored_usage)
        return public_turn(db.one("SELECT * FROM office_turns WHERE id=?", (run,)))


async def server_turn(prepared):
    run, key, user = prepared["run_id"], prepared["conversation_id"], prepared["user"]
    assignment = prepared["assignment"]
    try:
        for label in prepared["steps"][:-1]:
            yield event("step", label=label)
        yield event("step", label=prepared["steps"][-1])
        async with PLAN_SLOTS:
            if gateway.MODE == "demo":
                if prepared["scope"]:
                    decision, usage = await decide_slide(prepared["body"], prepared["scope"], prepared["sources"],
                        prepared["project"], assignment, user, run)
                else:
                    decision = demo_decision(prepared["body"].message, prepared["body"].mode,
                        prepared["task"], prepared["sources"], prepared["recent"], prepared["latest"])
                    usage = {}
                text = decision.model_dump_json()
            else:
                text, usage = await gateway.request_model(prepared["prompt"], run,
                    gateway.headers_for(user["id"], assignment, run), prepared["system"], prepared["context"])
        final = await complete_turn(key, run,
            TurnCompletion(text=text, usage=usage, model=gateway.MODEL if gateway.MODE != "demo" else "demo"),
            user, expected_engine=prepared["engine"])
        yield event("complete", turn=final)
    except asyncio.CancelledError:
        fail_turn(run, user, assignment, "本轮连接已断开，可以重新发送")
        raise
    except HTTPException as error:
        fail_turn(run, user, assignment, str(error.detail))
        yield event("error", message=str(error.detail))
    except Exception:
        LOG.exception("Office turn failed on run %s", run)
        message = "本轮未能完成，请重试。此前的会话和已保存文件仍可查阅。"
        fail_turn(run, user, assignment, message)
        yield event("error", message=message)
@router.post("/{key}/images/prepare")
def prepare_image_route(key: str, body: ImagePrepare, user: User):
    if gateway.MODE != "desktop":
        fail(409, "桌面执行模式未启用")
    run, now, stale = uid(), instant(), stale_before()
    with database(write=True) as db:
        session = owned_thread(db, key, user)
        assignment = session["assignment_id"]
        writable_task(db, assignment, user)
        for prior in db.all("SELECT id,result FROM office_turns WHERE conversation_id=? AND status='running' AND created_at<?",
                            (key, stale)):
            result = unpack(prior["result"] or "{}")
            result["error"] = "上次执行已中断，请重新发送"
            db.execute("UPDATE office_turns SET status='failed',result=? WHERE id=? AND status='running'",
                       (pack(result), prior["id"]))
            db.execute("UPDATE runs SET status='failed' WHERE id=? AND status='running'", (prior["id"],))
        if db.one("SELECT id FROM office_turns WHERE conversation_id=? AND status='running'", (key,)):
            fail(409, "当前会话有一轮任务正在执行")
        image_request = {"prompt": body.prompt, "model": body.model, "size": body.size, "quality": body.quality}
        request = pack({"message": "生成图片：" + body.prompt, "kind": "image", "image": image_request})
        result = {"kind": "image", "prompt": body.prompt, "model": body.model, "size": body.size,
                  "quality": body.quality, "images": []}
        db.execute("INSERT INTO runs(id,assignment_id,user_id,kind,engine,status,usage,created_at) VALUES(?,?,?,?,?,?,?,?)",
                   (run, assignment, user["id"], "image", "desktop", "running", pack({"model": body.model}), now))
        db.execute("INSERT INTO office_turns(id,conversation_id,request_id,input,result,status,created_at) VALUES(?,?,?,?,?,?,?)",
                   (run, key, run, request, pack(result), "running", now))
        db.execute("UPDATE office_conversations SET updated_at=? WHERE id=?", (now, key))
        audit(db, "image.started", user["id"], assignment, run_id=run, model=body.model,
              size=body.size, quality=body.quality)
    return {"run_id": run}


@router.post("/{key}/images/{run}/complete")
def complete_image_route(key: str, run: str, payload: ImageCompletion, user: User):
    if gateway.MODE != "desktop":
        fail(409, "桌面执行模式未启用")
    data = decode_image(payload)
    usage = image_usage(payload.usage)
    with database(write=True) as db:
        session = owned_thread(db, key, user)
        assignment = session["assignment_id"]
        writable_task(db, assignment, user)
        turn = db.one("SELECT * FROM office_turns WHERE id=? AND conversation_id=?", (run, key))
        run_row = db.one("SELECT * FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
        if not turn or not run_row or run_row["engine"] != "desktop" or run_row["kind"] != "image":
            fail(404, "生图执行不存在")
        if turn["status"] != "running" or run_row["status"] != "running":
            fail(409, "本轮生图已结束")
        original = unpack(turn["input"]).get("image", {})
        if any(payload_value != original.get(name) for name, payload_value in
               (("model", payload.model), ("size", payload.size), ("quality", payload.quality))):
            fail(422, "生图参数与准备请求不一致")
        result = unpack(turn["result"] or "{}")
        index = len(result.get("images", [])) + 1
        name = safe_name(f"生成图片-{index}") + ".png"
        file_id = add_file(db, user["id"], name, data, "png", "image", assignment=assignment, run=run)
        image = {"file": public_file(db, file_id), "prompt": original["prompt"], "model": payload.model,
                 "size": payload.size, "quality": payload.quality, "width": payload.width,
                 "height": payload.height, "source": "user", "usage": usage}
        result.setdefault("images", []).append(image)
        result.update(kind="image", prompt=original["prompt"], model=payload.model, size=payload.size,
                      quality=payload.quality, width=payload.width, height=payload.height)
        changed = db.execute("UPDATE office_turns SET result=?,status='completed' WHERE id=? AND status='running'",
                             (pack(result), run)).rowcount
        if not changed:
            fail(409, "本轮生图已结束")
        db.execute("UPDATE runs SET status='completed',usage=? WHERE id=? AND status='running'",
                   (pack({"model": payload.model, **usage}), run))
        db.execute("UPDATE office_conversations SET updated_at=? WHERE id=?", (instant(), key))
        audit(db, "image.generated", user["id"], assignment, run_id=run, artifact_id=file_id,
              model=payload.model, size=payload.size, quality=payload.quality, usage=usage)
        return public_turn(db.one("SELECT * FROM office_turns WHERE id=?", (run,)))


@router.post("/{key}/images/{run}/fail")
def fail_image_route(key: str, run: str, payload: ImageFailure, user: User):
    with database() as db:
        assignment = owned_thread(db, key, user)["assignment_id"]
        turn = db.one("SELECT status FROM office_turns WHERE id=? AND conversation_id=?", (run, key))
        run_row = db.one("SELECT status,engine,kind FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
        if not turn or not run_row or run_row["engine"] != "desktop" or run_row["kind"] != "image":
            fail(404, "生图执行不存在")
        if turn["status"] != "running" or run_row["status"] != "running":
            fail(409, "本轮生图已结束")
    if not fail_turn(run, user, assignment, clean_text(payload.message, 300)):
        fail(409, "本轮生图已结束")
    return {"ok": True}


@router.post("/{key}/turns/{run}/images")
def store_turn_image_route(key: str, run: str, payload: TurnImage, user: User):
    if gateway.MODE != "desktop":
        fail(409, "桌面执行模式未启用")
    data = decode_image(payload)
    usage = image_usage(payload.usage)
    with database(write=True) as db:
        session = owned_thread(db, key, user)
        assignment = session["assignment_id"]
        writable_task(db, assignment, user)
        turn = db.one("SELECT * FROM office_turns WHERE id=? AND conversation_id=?", (run, key))
        run_row = db.one("SELECT * FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
        if (not turn or not run_row or run_row["engine"] != "desktop" or run_row["kind"] != "conversation"):
            fail(404, "本轮执行不存在")
        if turn["status"] != "running" or run_row["status"] != "running":
            fail(409, "本轮执行已结束")
        result = unpack(turn["result"] or "{}")
        images = result.setdefault("images", [])
        index = len(images) + 1
        file_id = add_file(db, user["id"], safe_name(f"AI配图-{index}") + ".png", data, "png", "image",
                           assignment=assignment, run=run)
        image = {"file": public_file(db, file_id), "prompt": payload.prompt, "model": payload.model,
                 "size": payload.size, "quality": payload.quality, "width": payload.width,
                 "height": payload.height, "source": "ai", "usage": usage}
        images.append(image)
        changed = db.execute("UPDATE office_turns SET result=? WHERE id=? AND status='running'",
                             (pack(result), run)).rowcount
        if not changed:
            fail(409, "本轮执行已结束")
        db.execute("UPDATE office_conversations SET updated_at=? WHERE id=?", (instant(), key))
        audit(db, "image.ai_generated", user["id"], assignment, run_id=run, artifact_id=file_id,
              model=payload.model, size=payload.size, quality=payload.quality, usage=usage)
        return image


@router.post("/{key}/turns/prepare")
async def prepare_turn_route(key: str, body: TurnRequest, user: User):
    if gateway.MODE != "desktop":
        fail(409, "桌面执行模式未启用")
    prepared = await prepare_conversation(key, body, user, engine="desktop")
    if "replay" in prepared:
        fail(409, "本轮请求已经提交，请等待执行完成")
    return {name: prepared[name] for name in ("run_id", "kind", "system", "prompt", "context", "steps", "wants_pptx", "image_index_offset", "web_search")}


@router.post("/{key}/turns/{run}/complete")
async def complete_turn_route(key: str, run: str, payload: TurnCompletion, user: User):
    if gateway.MODE != "desktop":
        fail(409, "桌面执行模式未启用")
    try:
        return await complete_turn(key, run, payload, user, expected_engine="desktop")
    except HTTPException as error:
        with database() as db:
            conversation = db.one("SELECT id FROM office_conversations WHERE id=? AND owner_id=?", (key, user["id"]))
            run_row = db.one("SELECT status,engine,assignment_id FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
            turn = db.one("SELECT status FROM office_turns WHERE id=? AND conversation_id=?", (run, key))
        if conversation and run_row and turn and run_row["engine"] == "desktop" and run_row["status"] == "running" and turn["status"] == "running":
            fail_turn(run, user, run_row["assignment_id"], str(error.detail))
        raise
    except Exception:
        with database() as db:
            conversation = db.one("SELECT id FROM office_conversations WHERE id=? AND owner_id=?", (key, user["id"]))
            run_row = db.one("SELECT status,engine,assignment_id FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
            turn = db.one("SELECT status FROM office_turns WHERE id=? AND conversation_id=?", (run, key))
        if conversation and run_row and turn and run_row["engine"] == "desktop" and run_row["status"] == "running" and turn["status"] == "running":
            fail_turn(run, user, run_row["assignment_id"], "本轮未能完成，请重试")
        raise


@router.post("/{key}/turns/{run}/fail")
def fail_turn_route(key: str, run: str, payload: TurnFailure, user: User):
    with database() as db:
        assignment = owned_thread(db, key, user)["assignment_id"]
        turn = db.one("SELECT status FROM office_turns WHERE id=? AND conversation_id=?", (run, key))
        run_row = db.one("SELECT status,engine FROM runs WHERE id=? AND user_id=?", (run, user["id"]))
        if not turn or not run_row or run_row["engine"] != "desktop":
            fail(404, "本轮执行不存在")
        if turn["status"] != "running" or run_row["status"] != "running":
            fail(409, "本轮执行已结束")
    if not fail_turn(run, user, assignment, clean_text(payload.message, 1000), steps=payload.steps):
        fail(409, "本轮执行已结束")
    return {"ok": True}


@router.post("/{key}/turns")
async def send_turn(key: str, body: TurnRequest, user: User):
    if gateway.MODE == "desktop":
        fail(409, "请在穹驭办公桌面客户端中使用 AI 办理")
    prepared = await prepare_conversation(key, body, user, engine=gateway.MODE, allow_replay=True)
    if "replay" in prepared:
        return stream(replay(prepared["replay"]))
    return stream(server_turn(prepared))
