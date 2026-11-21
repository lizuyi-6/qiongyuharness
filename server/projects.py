import asyncio
import os
import re
import secrets
from typing import Annotated, Any

from fastapi import APIRouter, File, Form, UploadFile
from pydantic import BaseModel, Field

from .identity import fail
from .office import MAX_FILE, add_file, clean_text, extract, verify
from .store import audit, database, instant, pack, uid, unpack
from .tasks import User

router = APIRouter(prefix="/api/v1/projects")
SCANNER = os.getenv("QO_CLAMD", "")

PROJECT_EXTENSIONS = frozenset({"docx", "pptx", "xlsx", "pdf", "txt", "md", "csv"})

CATEGORIES: dict[str, str] = {
    "goal": "背景与目标",
    "decision": "纪要与决策",
    "rule": "口径与规范",
    "doc": "方案与资料",
    "data": "数据与报表",
}

_HEADING_RE = re.compile(
    r"^(?:#{1,4}\s+\S|[一二三四五六七八九十百]+[、.．]\s*\S|（[一二三四五六七八九十]+）\s*\S|"
    r"\([一二三四五六七八九十\d]+\)\s*\S|\d{1,2}[.、．]\s*\S|第[一二三四五六七八九十\d]+[章节部分条篇]\s*\S)"
)
_SLIDE_MARKER_RE = re.compile(r"^【第\s*(\d+)\s*页】$")
_SHEET_MARKER_RE = re.compile(r"^【工作表\s*(\d+)】$")
_HIGH_SIGNAL_RE = re.compile(
    r"(?:结论|决定|决议|目标|要求|重点|核心|关键|截止|交付|负责|验收|风险|注意|必须|严禁|口径|指标|增长|完成|预算|排期|"
    r"\d+(?:\.\d+)?%|\d{1,2}月\d{1,2}日|周[一二三四五六日天]|\d+(?:\.\d+)?(?:万|亿|元|人|页|天|小时))"
)


class ProjectBody(BaseModel):
    name: str = Field(min_length=1, max_length=80)
    description: str = Field(default="", max_length=2000)


class ProjectUpdateBody(BaseModel):
    name: str | None = Field(default=None, max_length=80)
    description: str | None = Field(default=None, max_length=2000)
    brief: str | None = Field(default=None, max_length=12000)


class JoinBody(BaseModel):
    code: str = Field(min_length=24, max_length=24, pattern=r"^[A-Za-z0-9_-]+$")


class ContextBody(BaseModel):
    title: str = Field(default="", max_length=110)
    content: str = Field(default="", max_length=8000)
    artifact_id: str | None = Field(default=None, max_length=100)
    category: str = Field(default="auto", max_length=24)
    pinned: bool = False


class ContextPatchBody(BaseModel):
    title: str | None = Field(default=None, max_length=110)
    content: str | None = Field(default=None, max_length=8000)
    category: str | None = Field(default=None, max_length=24)
    pinned: bool | None = None
    enabled: bool | None = None


class BatchFilesBody(BaseModel):
    artifact_ids: list[str] = Field(min_length=1, max_length=16)
    category: str = Field(default="auto", max_length=24)
    pinned: bool = False


def project_for(db, project_id, user):
    row = db.one(
        """SELECT p.* FROM office_projects p JOIN office_project_members m ON m.project_id=p.id
           WHERE p.id=? AND m.user_id=?""",
        (project_id, user["id"]),
    )
    if not row:
        fail(404, "项目不存在或当前账号无权访问")
    return row


def shared_file_access(db, file_id, user_id):
    return (
        db.one(
            """SELECT c.id FROM office_project_context c JOIN office_project_members m ON m.project_id=c.project_id
               WHERE c.artifact_id=? AND m.user_id=? LIMIT 1""",
            (file_id, user_id),
        )
        is not None
    )


def infer_category(title: str = "", content: str = "", filename: str = "", ext: str = "", requested: str = "") -> str:
    req = (requested or "").strip().lower()
    if req in CATEGORIES:
        return req
    low_ext = (ext or "").strip().lower()
    probe = f"{title} {filename} {(content or '')[:500]}"
    if re.search(r"纪要|会议|决议|决定|拍板|批示|评审结论|例会|待办|行动项|复盘", probe):
        return "decision"
    if re.search(r"规范|口径|标准|模板|指南|要求|红线|准则|术语|制度|手册", probe):
        return "rule"
    if re.search(r"立项|背景|目标|规划|需求|愿景|OKR|KPI|使命|范围|章程", probe):
        return "goal"
    if low_ext in {"xlsx", "csv"} or re.search(r"数据|报表|清单|台账|统计|预算|明细|汇总表|指标表", probe):
        return "data"
    return "doc"


def _strip_heading_prefix(line: str) -> str:
    cleaned = re.sub(r"^#{1,6}\s*", "", line.strip())
    return " ".join(cleaned.split())[:48]


def preorganize_text(
    title: str,
    note_content: str = "",
    extracted_text: str = "",
    filename: str = "",
    ext: str = "",
    requested_category: str = "",
) -> dict[str, Any]:
    low_ext = (ext or "").strip().lower()
    category = infer_category(title, note_content or extracted_text, filename, low_ext, requested_category)
    raw_source = "\n".join(part for part in (note_content.strip(), extracted_text.strip()) if part).strip()
    lines = [" ".join(line.split()) for line in raw_source.splitlines() if line.strip()]
    total_chars = len(raw_source)

    outline: list[str] = []
    seen_outline: set[str] = set()

    def add_outline(item: str):
        norm = _strip_heading_prefix(item)
        if norm and len(norm) >= 2 and norm not in seen_outline and len(outline) < 8:
            seen_outline.add(norm)
            outline.append(norm)

    # 1. 计算结构规模与大纲
    unit_label = "文本笔记"
    if low_ext == "pptx":
        slide_count = 0
        for idx, line in enumerate(lines):
            m = _SLIDE_MARKER_RE.match(line)
            if m:
                slide_count += 1
                if idx + 1 < len(lines) and not _SLIDE_MARKER_RE.match(lines[idx + 1]):
                    add_outline(f"P{m.group(1)} {lines[idx + 1][:32]}")
        unit_label = f"{max(1, slide_count)} 页幻灯片"
    elif low_ext == "xlsx":
        sheet_count = 0
        data_rows = 0
        for idx, line in enumerate(lines):
            m = _SHEET_MARKER_RE.match(line)
            if m:
                sheet_count += 1
                if idx + 1 < len(lines) and not _SHEET_MARKER_RE.match(lines[idx + 1]):
                    cols = [c.strip() for c in lines[idx + 1].split("|") if c.strip()][:5]
                    if cols:
                        add_outline(f"表{m.group(1)}：{' / '.join(cols)[:38]}")
            else:
                data_rows += 1
        unit_label = f"{max(1, sheet_count)} 张工作表 · {data_rows} 行数据"
    elif low_ext == "csv":
        data_rows = len(lines)
        if lines:
            cols = [c.strip() for c in re.split(r"[,;\t|]", lines[0]) if c.strip()][:6]
            if cols:
                add_outline(f"字段：{' / '.join(cols)[:42]}")
        unit_label = f"{data_rows} 行表格数据"
    else:
        for line in lines:
            if _SLIDE_MARKER_RE.match(line) or _SHEET_MARKER_RE.match(line):
                continue
            if _HEADING_RE.match(line):
                add_outline(line)
            elif 2 <= len(line) <= 24 and not line.endswith(("。", "；", "，", ".", ";", ",", "！", "!", "？", "?")):
                if re.search(r"[:：]$", line) or len(outline) < 4:
                    add_outline(line.rstrip(":："))
        if low_ext == "docx":
            unit_label = f"Word 文档 · {max(1, len(lines))} 段"
        elif low_ext == "pdf":
            unit_label = f"PDF 文档 · {max(1, len(lines))} 段"
        elif low_ext in {"md", "txt"}:
            unit_label = f"{low_ext.upper()} 文本 · {max(1, len(lines))} 行"
        elif lines:
            unit_label = f"上下文纪要 · {len(lines)} 条"

    # 2. 提炼核心要点与结论
    candidate_points: list[tuple[int, int, str]] = []
    seen_points: set[str] = set()
    for idx, line in enumerate(lines):
        if _SLIDE_MARKER_RE.match(line) or _SHEET_MARKER_RE.match(line):
            continue
        cleaned = re.sub(r"^(?:[-*•·]|\d+[.、．])\s*", "", _strip_heading_prefix(line)).strip()
        if len(cleaned) < 4 or cleaned in seen_outline:
            continue
        # 切分长段落为独立句子
        segments = [seg.strip() for seg in re.split(r"(?<=[。！？；!?;])\s*", cleaned) if len(seg.strip()) >= 4]
        for seg in (segments or [cleaned])[:3]:
            snippet = seg[:110]
            if snippet in seen_points:
                continue
            seen_points.add(snippet)
            score = 1
            if _HIGH_SIGNAL_RE.search(snippet):
                score += 3
            if idx < 5:
                score += 2
            if note_content and snippet in note_content:
                score += 2
            candidate_points.append((score, -idx, snippet))

    candidate_points.sort(key=lambda item: (item[0], item[1]), reverse=True)
    key_points = [item[2] for item in candidate_points[:5]]
    if not key_points and lines:
        first_valid = next((l for l in lines if not _SLIDE_MARKER_RE.match(l) and not _SHEET_MARKER_RE.match(l)), "")
        if first_valid:
            key_points = [first_valid[:110]]

    # 3. 组装一句话预整理摘要
    summary_parts: list[str] = []
    if note_content.strip():
        note_one_line = " ".join(note_content.split())
        summary_parts.append(note_one_line[:110])
    elif key_points:
        summary_parts.append("；".join(key_points[:2])[:130])
    if outline and (not note_content.strip() or bool(extracted_text.strip())):
        summary_parts.append("结构：" + "、".join(outline[:4])[:80])
    summary = " ｜ ".join(summary_parts)[:220] if summary_parts else (f"已解析 {filename or title}" if total_chars else "已归档到项目上下文")

    excerpt = "\n".join(
        l for l in lines if not _SLIDE_MARKER_RE.match(l) and not _SHEET_MARKER_RE.match(l)
    )[:1600]

    return {
        "category": category,
        "category_label": CATEGORIES.get(category, "方案与资料"),
        "stats": {
            "chars": total_chars,
            "unit_label": unit_label,
        },
        "outline": outline[:8],
        "key_points": key_points[:5],
        "summary": summary,
        "excerpt": excerpt,
    }


def preorganize_item_with_bytes(
    title: str,
    note_content: str = "",
    filename: str = "",
    ext: str = "",
    file_bytes: bytes | None = None,
    requested_category: str = "",
) -> dict[str, Any]:
    extracted = ""
    if file_bytes and ext in PROJECT_EXTENSIONS:
        try:
            extracted = extract(file_bytes, ext, limit=20000) or ""
        except Exception:
            extracted = ""
    return preorganize_text(
        title=title,
        note_content=note_content,
        extracted_text=extracted,
        filename=filename,
        ext=ext,
        requested_category=requested_category,
    )


def _parse_item_digest(row: Any) -> dict[str, Any]:
    keys = row.keys() if hasattr(row, "keys") else ()
    raw_digest = row["digest"] if "digest" in keys else ""
    cat = (row["category"] if "category" in keys else "") or ""
    filename = (row["filename"] if "filename" in keys else "") or ""
    ext = (row["extension"] if "extension" in keys else "") or ""
    if raw_digest:
        try:
            parsed = unpack(raw_digest)
            if isinstance(parsed, dict) and parsed.get("stats"):
                resolved_cat = cat if cat in CATEGORIES else parsed.get("category") or "doc"
                parsed["category"] = resolved_cat
                parsed["category_label"] = CATEGORIES.get(resolved_cat, "方案与资料")
                return parsed
        except Exception:
            pass
    return preorganize_text(
        title=row["title"] or filename or "项目资料",
        note_content=row["content"] or "",
        extracted_text="",
        filename=filename,
        ext=ext,
        requested_category=cat,
    )


def format_context_row(row: Any) -> dict[str, Any]:
    keys = row.keys() if hasattr(row, "keys") else ()
    digest = _parse_item_digest(row)
    category = digest["category"]
    return {
        "id": row["id"],
        "title": row["title"],
        "content": row["content"] or "",
        "category": category,
        "category_label": CATEGORIES.get(category, "方案与资料"),
        "pinned": bool(row["pinned"]) if "pinned" in keys else False,
        "enabled": bool(row["enabled"]) if "enabled" in keys and row["enabled"] is not None else True,
        "artifact_id": row["artifact_id"],
        "filename": row["filename"] if "filename" in keys else None,
        "extension": row["extension"] if "extension" in keys else None,
        "size": row["size"] if "size" in keys else None,
        "author": row["author"] if "author" in keys else "",
        "author_id": row["author_id"] if "author_id" in keys else "",
        "created_at": row["created_at"],
        "updated_at": (row["updated_at"] if "updated_at" in keys and row["updated_at"] else row["created_at"]),
        "digest": digest,
    }


def synthesize_auto_brief_text(project_row: Any, items: list[dict[str, Any]]) -> str:
    keys = project_row.keys() if hasattr(project_row, "keys") else ()
    description = (project_row["description"] if "description" in keys else "") or ""
    enabled_items = [it for it in items if it.get("enabled", True)]

    sections: list[str] = []
    goal_lines: list[str] = []
    if description.strip():
        goal_lines.append(f"- 项目目标说明：{description.strip()}")
    for it in enabled_items:
        if it["category"] == "goal":
            pt = it["content"].strip() or (it["digest"].get("summary") or "")
            goal_lines.append(f"- 【{it['title']}】{pt[:180]}")
    if goal_lines:
        sections.append("### 一、项目背景与目标\n" + "\n".join(goal_lines[:8]))

    decision_lines: list[str] = []
    for it in enabled_items:
        if it["category"] in {"decision", "rule"} or it.get("pinned"):
            tag = "置顶" if it.get("pinned") else it["category_label"]
            pt = it["content"].strip() or "；".join(it["digest"].get("key_points") or []) or (it["digest"].get("summary") or "")
            if pt:
                decision_lines.append(f"- [{tag}] 【{it['title']}】：{pt[:200]}")
    if decision_lines:
        sections.append("### 二、核心决策与执行口径\n" + "\n".join(decision_lines[:10]))

    doc_lines: list[str] = []
    for it in enabled_items:
        d = it.get("digest") or {}
        stats = d.get("stats") or {}
        outline = d.get("outline") or []
        kps = d.get("key_points") or []
        file_tag = f"（文件：{it['filename']} · {stats.get('unit_label', '')}）" if it.get("filename") else ""
        outline_str = f" | 结构：{'、'.join(outline[:4])}" if outline else ""
        kp_str = f" | 要点：{'；'.join(kps[:2])[:120]}" if kps else (f" | 说明：{it['content'][:120]}" if it.get("content") else "")
        doc_lines.append(f"- [{it['category_label']}] **{it['title']}**{file_tag}{outline_str}{kp_str}")
    if doc_lines:
        sections.append("### 三、已预整理资料结构索引\n" + "\n".join(doc_lines[:15]))

    return "\n\n".join(sections).strip()


def build_project_overview(project_row: Any, items: list[dict[str, Any]]) -> dict[str, Any]:
    keys = project_row.keys() if hasattr(project_row, "keys") else ()
    custom_brief = (project_row["brief"] if "brief" in keys else "") or ""
    description = (project_row["description"] if "description" in keys else "") or ""
    enabled_items = [it for it in items if it.get("enabled", True)]
    pinned_items = [it for it in enabled_items if it.get("pinned")]
    file_items = [it for it in items if it.get("artifact_id")]
    total_chars = sum(int(((it.get("digest") or {}).get("stats") or {}).get("chars") or len(it.get("content") or "")) for it in enabled_items)
    if description:
        total_chars += len(description)

    cat_counts: dict[str, int] = {k: 0 for k in CATEGORIES}
    for it in items:
        c = it.get("category") or "doc"
        cat_counts[c] = cat_counts.get(c, 0) + 1

    goals: list[str] = []
    if description.strip():
        goals.append(description.strip())
    for it in enabled_items:
        if it["category"] == "goal":
            text = it["content"].strip() or (it["digest"].get("summary") or "")
            if text:
                goals.append(f"{it['title']}：{text[:160]}")

    decisions: list[str] = []
    for it in enabled_items:
        if it["category"] in {"decision", "rule"} or it.get("pinned"):
            text = it["content"].strip() or "；".join(it["digest"].get("key_points") or []) or (it["digest"].get("summary") or "")
            if text:
                decisions.append(f"【{it['title']}】{text[:180]}")

    auto_brief = synthesize_auto_brief_text(project_row, items)
    return {
        "custom_brief": custom_brief,
        "auto_brief": auto_brief,
        "effective_brief": custom_brief.strip() or auto_brief,
        "goals": goals[:6],
        "decisions": decisions[:8],
        "stats": {
            "total_count": len(items),
            "enabled_count": len(enabled_items),
            "pinned_count": len(pinned_items),
            "file_count": len(file_items),
            "note_count": len(items) - len(file_items),
            "total_chars": total_chars,
            "category_counts": cat_counts,
        },
    }


def summary(row, user, extra: dict[str, Any] | None = None):
    keys = row.keys() if hasattr(row, "keys") else ()
    base = {
        "id": row["id"],
        "name": row["name"],
        "description": (row["description"] if "description" in keys else "") or "",
        "has_brief": bool((row["brief"] if "brief" in keys else "") or ""),
        "is_owner": row["owner_id"] == user["id"],
        "created_at": row["created_at"] if "created_at" in keys else "",
        "updated_at": (row["updated_at"] if "updated_at" in keys and row["updated_at"] else (row["created_at"] if "created_at" in keys else "")),
    }
    if extra:
        base.update(extra)
    return base


def _load_project_context_items(db, project_id: str, limit: int = 80) -> list[dict[str, Any]]:
    rows = db.all(
        """SELECT c.id, c.title, c.content, c.category, c.digest, c.pinned, c.enabled,
                  c.artifact_id, c.author_id, c.created_at, c.updated_at,
                  u.name AS author, a.name AS filename, a.extension AS extension, a.size AS size
           FROM office_project_context c
           JOIN users u ON u.id=c.author_id
           LEFT JOIN artifacts a ON a.id=c.artifact_id
           WHERE c.project_id=?
           ORDER BY COALESCE(c.pinned, 0) DESC, c.created_at DESC LIMIT ?""",
        (project_id, limit),
    )
    return [format_context_row(r) for r in rows]


@router.get("")
def list_projects(user: User):
    with database() as db:
        rows = db.all(
            """SELECT p.*,
                      (SELECT COUNT(*) FROM office_project_members pm WHERE pm.project_id=p.id) AS member_count,
                      (SELECT COUNT(*) FROM office_project_context pc WHERE pc.project_id=p.id) AS context_count,
                      (SELECT COUNT(*) FROM office_project_context pc WHERE pc.project_id=p.id AND COALESCE(pc.enabled,1)=1) AS enabled_context_count,
                      (SELECT COUNT(*) FROM office_project_context pc WHERE pc.project_id=p.id AND pc.artifact_id IS NOT NULL) AS file_count
               FROM office_projects p
               JOIN office_project_members m ON m.project_id=p.id
               WHERE m.user_id=?
               ORDER BY COALESCE(NULLIF(p.updated_at,''), m.joined_at) DESC LIMIT 60""",
            (user["id"],),
        )
        return [
            summary(
                row,
                user,
                {
                    "member_count": int(row["member_count"] or 1),
                    "context_count": int(row["context_count"] or 0),
                    "enabled_context_count": int(row["enabled_context_count"] or 0),
                    "file_count": int(row["file_count"] or 0),
                },
            )
            for row in rows
        ]


@router.post("")
def create_project(body: ProjectBody, user: User):
    key, now = uid(), instant()
    name = clean_text(body.name, 80)
    description = clean_text(body.description, 2000, False)
    with database(write=True) as db:
        db.execute(
            "INSERT INTO office_projects(id,owner_id,name,description,brief,invite_code,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)",
            (key, user["id"], name, description, "", secrets.token_urlsafe(18), now, now),
        )
        db.execute(
            "INSERT INTO office_project_members(project_id,user_id,joined_at) VALUES(?,?,?)",
            (key, user["id"], now),
        )
        audit(db, "project.created", user["id"], project_id=key)
        return summary(project_for(db, key, user), user, {"member_count": 1, "context_count": 0, "enabled_context_count": 0, "file_count": 0})


@router.post("/join")
def join_project(body: JoinBody, user: User):
    with database(write=True) as db:
        row = db.one("SELECT * FROM office_projects WHERE invite_code=?", (body.code,))
        if not row:
            fail(404, "邀请码无效")
        changed = db.execute(
            """INSERT INTO office_project_members(project_id,user_id,joined_at) VALUES(?,?,?)
               ON CONFLICT(project_id,user_id) DO NOTHING""",
            (row["id"], user["id"], instant()),
        ).rowcount
        if changed:
            audit(db, "project.joined", user["id"], project_id=row["id"])
        return summary(row, user)


@router.get("/{key}")
def project_detail(key: str, user: User):
    with database() as db:
        row = project_for(db, key, user)
        members = db.all(
            """SELECT u.name, u.id, u.department, u.role FROM office_project_members m JOIN users u ON u.id=m.user_id
               WHERE m.project_id=? ORDER BY m.joined_at LIMIT 120""",
            (key,),
        )
        items = _load_project_context_items(db, key, limit=80)
        overview = build_project_overview(row, items)
        conversations = db.all(
            """SELECT id, title, updated_at, pinned FROM office_conversations
               WHERE project_id=? AND owner_id=?
               ORDER BY pinned DESC, updated_at DESC LIMIT 20""",
            (key, user["id"]),
        )
        return {
            **summary(row, user),
            "brief": (row["brief"] if "brief" in row.keys() else "") or "",
            "invite_code": row["invite_code"] if row["owner_id"] == user["id"] else None,
            "members": [dict(m) for m in members],
            "context": items,
            "overview": overview,
            "conversations": [dict(c) for c in conversations],
            "categories": CATEGORIES,
        }


@router.patch("/{key}")
def update_project(key: str, body: ProjectUpdateBody, user: User):
    now = instant()
    with database(write=True) as db:
        row = project_for(db, key, user)
        keys = row.keys() if hasattr(row, "keys") else ()
        next_name = row["name"]
        if body.name is not None:
            if row["owner_id"] != user["id"] and body.name.strip() != row["name"]:
                fail(403, "只有项目创建者可以修改项目名称")
            next_name = clean_text(body.name, 80)
        next_desc = (row["description"] if "description" in keys else "") or ""
        if body.description is not None:
            next_desc = clean_text(body.description, 2000, False)
        next_brief = (row["brief"] if "brief" in keys else "") or ""
        if body.brief is not None:
            next_brief = clean_text(body.brief, 12000, False)
        db.execute(
            "UPDATE office_projects SET name=?, description=?, brief=?, updated_at=? WHERE id=?",
            (next_name, next_desc, next_brief, now, key),
        )
        audit(db, "project.updated", user["id"], project_id=key)
    return project_detail(key, user)


@router.post("/{key}/invite")
def renew_invite(key: str, user: User):
    with database(write=True) as db:
        row = project_for(db, key, user)
        if row["owner_id"] != user["id"]:
            fail(403, "只有项目创建者可以更新邀请码")
        code = secrets.token_urlsafe(18)
        db.execute("UPDATE office_projects SET invite_code=?, updated_at=? WHERE id=?", (code, instant(), key))
        return {"invite_code": code}


async def _read_upload_bytes(file: UploadFile) -> bytes:
    try:
        return await file.read(MAX_FILE + 1)
    finally:
        await file.close()


@router.post("/{key}/files")
async def upload_project_file(
    key: str,
    user: User,
    file: Annotated[UploadFile, File()],
    title: Annotated[str, Form()] = "",
    content: Annotated[str, Form()] = "",
    category: Annotated[str, Form()] = "auto",
    pinned: Annotated[str, Form()] = "0",
):
    with database() as db:
        project_for(db, key, user)
    data = await _read_upload_bytes(file)
    name, ext = await asyncio.to_thread(verify, file.filename or "", data, SCANNER)
    if ext not in PROJECT_EXTENSIONS:
        fail(422, "项目知识库支持 Word、PPT、Excel、PDF、TXT、Markdown 和 CSV 文档")
    stem = name.rsplit(".", 1)[0].strip() if "." in name else name.strip()
    clean_title = clean_text(title.strip() or stem or name, 110)
    clean_content = clean_text(content, 8000, False)
    is_pinned = 1 if str(pinned).strip().lower() in {"1", "true", "yes"} else 0
    digest = await asyncio.to_thread(
        preorganize_item_with_bytes,
        clean_title,
        clean_content,
        name,
        ext,
        data,
        category,
    )
    resolved_cat = digest["category"]
    now = instant()
    with database(write=True) as db:
        project_for(db, key, user)
        artifact_id = add_file(db, user["id"], name, data, ext, "uploaded")
        entry = uid()
        db.execute(
            """INSERT INTO office_project_context(
                 id,project_id,author_id,title,content,category,digest,pinned,enabled,artifact_id,created_at,updated_at
               ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (entry, key, user["id"], clean_title, clean_content, resolved_cat, pack(digest), is_pinned, 1, artifact_id, now, now),
        )
        db.execute("UPDATE office_projects SET updated_at=? WHERE id=?", (now, key))
        audit(db, "project.file_uploaded", user["id"], project_id=key, context_id=entry, artifact_id=artifact_id)
        created_row = db.one(
            """SELECT c.id, c.title, c.content, c.category, c.digest, c.pinned, c.enabled,
                      c.artifact_id, c.author_id, c.created_at, c.updated_at,
                      u.name AS author, a.name AS filename, a.extension AS extension, a.size AS size
               FROM office_project_context c JOIN users u ON u.id=c.author_id
               LEFT JOIN artifacts a ON a.id=c.artifact_id WHERE c.id=?""",
            (entry,),
        )
        return format_context_row(created_row)


@router.post("/{key}/context")
def publish_context(key: str, body: ContextBody, user: User):
    content = clean_text(body.content, 8000, False)
    raw_title = (body.title or "").strip()
    if not content and not body.artifact_id:
        fail(422, "请填写项目资料正文或选择文件")
    now = instant()
    with database(write=True) as db:
        project_for(db, key, user)
        file_row = None
        if body.artifact_id:
            file_row = db.one(
                "SELECT id,owner_id,assignment_id,label,name,extension,data,size FROM artifacts WHERE id=?",
                (body.artifact_id,),
            )
            if not file_row or file_row["owner_id"] != user["id"] or file_row["assignment_id"] or file_row["label"] == "material":
                fail(403, "只能共享自己个人工作区里的文件")
            if file_row["extension"] not in PROJECT_EXTENSIONS:
                fail(422, "项目上下文支持 Word、PPT、Excel、PDF、TXT、Markdown 和 CSV")
        stem = (file_row["name"].rsplit(".", 1)[0] if file_row and "." in file_row["name"] else (file_row["name"] if file_row else ""))
        fallback_title = stem or (content.splitlines()[0][:36] if content else "项目上下文资料")
        title = clean_text(raw_title or fallback_title, 110)
        digest = preorganize_item_with_bytes(
            title=title,
            note_content=content,
            filename=file_row["name"] if file_row else "",
            ext=file_row["extension"] if file_row else "",
            file_bytes=bytes(file_row["data"]) if file_row and file_row["data"] else None,
            requested_category=body.category,
        )
        resolved_cat = digest["category"]
        is_pinned = 1 if body.pinned else 0
        if body.artifact_id:
            prior = db.one(
                "SELECT id FROM office_project_context WHERE project_id=? AND artifact_id=?",
                (key, body.artifact_id),
            )
            if prior:
                db.execute(
                    """UPDATE office_project_context
                       SET title=?, content=CASE WHEN ?<>'' THEN ? ELSE content END,
                           category=?, digest=?, pinned=CASE WHEN ?=1 THEN 1 ELSE pinned END,
                           enabled=1, updated_at=?
                       WHERE id=?""",
                    (title, content, content, resolved_cat, pack(digest), is_pinned, now, prior["id"]),
                )
                db.execute("UPDATE office_projects SET updated_at=? WHERE id=?", (now, key))
                return {"id": prior["id"]}
        entry = uid()
        db.execute(
            """INSERT INTO office_project_context(
                 id,project_id,author_id,title,content,category,digest,pinned,enabled,artifact_id,created_at,updated_at
               ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (entry, key, user["id"], title, content, resolved_cat, pack(digest), is_pinned, 1, body.artifact_id, now, now),
        )
        db.execute("UPDATE office_projects SET updated_at=? WHERE id=?", (now, key))
        audit(db, "project.context_published", user["id"], project_id=key, context_id=entry)
        return {"id": entry}


@router.post("/{key}/context/batch-files")
def batch_share_files(key: str, body: BatchFilesBody, user: User):
    added: list[str] = []
    skipped: list[str] = []
    now = instant()
    with database(write=True) as db:
        project_for(db, key, user)
        for fid in dict.fromkeys(body.artifact_ids):
            fid = str(fid or "").strip()
            if not fid:
                continue
            file_row = db.one(
                "SELECT id,owner_id,assignment_id,label,name,extension,data FROM artifacts WHERE id=?",
                (fid,),
            )
            if not file_row or file_row["owner_id"] != user["id"] or file_row["assignment_id"] or file_row["label"] == "material":
                skipped.append(fid)
                continue
            if file_row["extension"] not in PROJECT_EXTENSIONS:
                skipped.append(fid)
                continue
            prior = db.one("SELECT id FROM office_project_context WHERE project_id=? AND artifact_id=?", (key, fid))
            if prior:
                db.execute("UPDATE office_project_context SET enabled=1, updated_at=? WHERE id=?", (now, prior["id"]))
                added.append(prior["id"])
                continue
            stem = file_row["name"].rsplit(".", 1)[0] if "." in file_row["name"] else file_row["name"]
            title = clean_text(stem or file_row["name"], 110)
            digest = preorganize_item_with_bytes(
                title=title,
                note_content="",
                filename=file_row["name"],
                ext=file_row["extension"],
                file_bytes=bytes(file_row["data"]),
                requested_category=body.category,
            )
            entry = uid()
            db.execute(
                """INSERT INTO office_project_context(
                     id,project_id,author_id,title,content,category,digest,pinned,enabled,artifact_id,created_at,updated_at
                   ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                (entry, key, user["id"], title, "", digest["category"], pack(digest), int(bool(body.pinned)), 1, fid, now, now),
            )
            audit(db, "project.context_published", user["id"], project_id=key, context_id=entry, artifact_id=fid)
            added.append(entry)
        if added:
            db.execute("UPDATE office_projects SET updated_at=? WHERE id=?", (now, key))
    return {"added": added, "skipped": skipped}


@router.patch("/{key}/context/{item_id}")
def update_context_item(key: str, item_id: str, body: ContextPatchBody, user: User):
    now = instant()
    with database(write=True) as db:
        project_for(db, key, user)
        row = db.one(
            """SELECT c.*, a.name AS filename, a.extension AS extension, a.data AS file_data
               FROM office_project_context c LEFT JOIN artifacts a ON a.id=c.artifact_id
               WHERE c.id=? AND c.project_id=?""",
            (item_id, key),
        )
        if not row:
            fail(404, "项目上下文条目不存在")
        keys = row.keys() if hasattr(row, "keys") else ()
        next_title = clean_text(body.title, 110) if body.title is not None else row["title"]
        next_content = clean_text(body.content, 8000, False) if body.content is not None else (row["content"] or "")
        if not next_content and not row["artifact_id"]:
            fail(422, "无附加文件时，上下文正文不能为空")
        cur_cat = (row["category"] if "category" in keys else "doc") or "doc"
        next_cat = body.category if (body.category in CATEGORIES or body.category == "auto") else cur_cat
        next_pinned = int(bool(body.pinned)) if body.pinned is not None else int(bool(row["pinned"] if "pinned" in keys else 0))
        next_enabled = int(bool(body.enabled)) if body.enabled is not None else int(bool(row["enabled"] if "enabled" in keys and row["enabled"] is not None else 1))

        if body.title is not None or body.content is not None or body.category is not None or not (row["digest"] if "digest" in keys else ""):
            digest = preorganize_item_with_bytes(
                title=next_title,
                note_content=next_content,
                filename=row["filename"] or "",
                ext=row["extension"] or "",
                file_bytes=bytes(row["file_data"]) if row["file_data"] else None,
                requested_category=next_cat,
            )
            resolved_cat = digest["category"]
            digest_json = pack(digest)
        else:
            resolved_cat = cur_cat
            digest_json = row["digest"]

        db.execute(
            """UPDATE office_project_context
               SET title=?, content=?, category=?, digest=?, pinned=?, enabled=?, updated_at=?
               WHERE id=? AND project_id=?""",
            (next_title, next_content, resolved_cat, digest_json, next_pinned, next_enabled, now, item_id, key),
        )
        db.execute("UPDATE office_projects SET updated_at=? WHERE id=?", (now, key))
        updated = db.one(
            """SELECT c.id, c.title, c.content, c.category, c.digest, c.pinned, c.enabled,
                      c.artifact_id, c.author_id, c.created_at, c.updated_at,
                      u.name AS author, a.name AS filename, a.extension AS extension, a.size AS size
               FROM office_project_context c JOIN users u ON u.id=c.author_id
               LEFT JOIN artifacts a ON a.id=c.artifact_id WHERE c.id=?""",
            (item_id,),
        )
        return format_context_row(updated)


@router.delete("/{key}/context/{item_id}")
def delete_context_item(key: str, item_id: str, user: User):
    with database(write=True) as db:
        proj = project_for(db, key, user)
        row = db.one(
            "SELECT id, author_id, title FROM office_project_context WHERE id=? AND project_id=?",
            (item_id, key),
        )
        if not row:
            fail(404, "项目上下文条目不存在")
        if row["author_id"] != user["id"] and proj["owner_id"] != user["id"]:
            fail(403, "仅发布人或项目创建者可移除该条共享资料")
        db.execute("DELETE FROM office_project_context WHERE id=? AND project_id=?", (item_id, key))
        db.execute("UPDATE office_projects SET updated_at=? WHERE id=?", (instant(), key))
        audit(db, "project.context_deleted", user["id"], project_id=key, context_id=item_id, title=row["title"])
    return {"ok": True, "id": item_id}


@router.post("/{key}/preorganize")
def preorganize_project(key: str, user: User):
    now = instant()
    with database(write=True) as db:
        proj = project_for(db, key, user)
        rows = db.all(
            """SELECT c.id, c.title, c.content, c.category, c.artifact_id,
                      a.name AS filename, a.extension AS extension, a.data AS file_data
               FROM office_project_context c
               LEFT JOIN artifacts a ON a.id=c.artifact_id
               WHERE c.project_id=?
               ORDER BY COALESCE(c.pinned, 0) DESC, c.created_at DESC""",
            (key,),
        )
        for r in rows:
            digest = preorganize_item_with_bytes(
                title=r["title"],
                note_content=r["content"] or "",
                filename=r["filename"] or "",
                ext=r["extension"] or "",
                file_bytes=bytes(r["file_data"]) if r["file_data"] else None,
                requested_category=r["category"] or "auto",
            )
            db.execute(
                "UPDATE office_project_context SET category=?, digest=?, updated_at=? WHERE id=?",
                (digest["category"], pack(digest), now, r["id"]),
            )
        items = _load_project_context_items(db, key, limit=80)
        auto_brief = synthesize_auto_brief_text(proj, items)
        db.execute("UPDATE office_projects SET brief=?, updated_at=? WHERE id=?", (auto_brief, now, key))
        audit(db, "project.preorganized", user["id"], project_id=key, count=len(items))

    detail = project_detail(key, user)
    ai_system = (
        "你是穹驭办公的项目知识库预整理专家。请根据给定的项目名称、目标说明与所有已启用的共享资料摘要，"
        "整理出一份结构清晰、高信息密度的《项目预整理上下文总览》（Markdown 格式，控制在 400–900 字），"
        "必须包含以下三个部分：\n"
        "### 一、项目背景与核心目标\n"
        "### 二、关键决策、约束与统一口径\n"
        "### 三、核心资料要点与待办索引\n"
        "直接输出 Markdown 正文，不要输出多余的寒暄或代码块包裹。"
    )
    ai_prompt = (
        f"项目名称：{detail['name']}\n"
        f"项目说明：{detail.get('description') or '（未填写）'}\n\n"
        f"当前已归档并初步预整理的资料清单：\n{detail['overview']['auto_brief'] or '（暂无资料）'}"
    )
    return {**detail, "ai_prepare": {"system": ai_system, "prompt": ai_prompt[:6000]}}


def project_sources(db, tag, user):
    if not tag:
        return None, [], []
    project = project_for(db, tag, user)
    pkeys = project.keys() if hasattr(project, "keys") else ()
    rows = db.all(
        """SELECT c.title, c.content, c.category, c.digest, c.pinned,
                  a.name, a.extension, a.data
           FROM office_project_context c
           LEFT JOIN artifacts a ON a.id=c.artifact_id
           WHERE c.project_id=? AND COALESCE(c.enabled, 1)=1
           ORDER BY COALESCE(c.pinned, 0) DESC, c.created_at DESC LIMIT 12""",
        (tag,),
    )
    notes = []
    files = []
    remaining = 14000
    for row in rows:
        if row["content"] and remaining > 0:
            excerpt = row["content"][: min(2600, remaining)]
            notes.append({"name": "项目：" + row["title"], "text": excerpt})
            remaining -= len(excerpt)
        if row["extension"] and row["data"] is not None:
            files.append(("项目：" + row["name"], bytes(row["data"]), row["extension"]))

    desc = (project["description"] if "description" in pkeys else "") or ""
    brief = (project["brief"] if "brief" in pkeys else "") or ""
    overview_parts = [p.strip() for p in (desc, brief) if p and p.strip()]
    if overview_parts and remaining > 200:
        overview_text = "\n\n".join(overview_parts)[: min(2800, remaining)]
        notes.append({"name": f"项目总览：{project['name']}", "text": overview_text})

    return project["name"], notes, files
