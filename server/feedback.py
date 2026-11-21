"""问题反馈与一键上传诊断日志接口。"""

from __future__ import annotations

import json
import re
from typing import Any
from urllib.parse import quote

from fastapi import APIRouter, Depends, Query
from fastapi.responses import Response
from pydantic import BaseModel, Field

from .identity import fail, user_for
from .store import audit, database, instant, uid, unpack

router = APIRouter(prefix="/api/v1/feedback")

CATEGORIES = {"bug", "ai", "ppt", "update", "suggestion", "other"}
CATEGORY_LABELS = {
    "bug": "功能异常",
    "ai": "AI 执行/生成失败",
    "ppt": "PPT 排版与画布",
    "update": "安装与更新",
    "suggestion": "体验建议",
    "other": "其他问题",
}
STATUSES = {"open", "in_progress", "resolved", "closed"}
STATUS_LABELS = {
    "open": "待处理",
    "in_progress": "处理中",
    "resolved": "已解决",
    "closed": "已关闭",
}
MAX_LOG_BYTES = 512 * 1024
APP_VERSION = "0.2.21"

_RE_BEARER = re.compile(r"\b(Bearer\s+)[A-Za-z0-9._~+/=-]{8,}", re.IGNORECASE)
_RE_SK_KEY = re.compile(r"\b(sk-[A-Za-z0-9_-]{8,})")
_RE_COOKIE = re.compile(r"\b(qo_session|qo_browser)=([^;\s\"'`]+)", re.IGNORECASE)
_RE_KV_SECRET = re.compile(
    r"((?:[\"'])?(?:api[_-]?key|access[_-]?token|refresh[_-]?token|id[_-]?token|client[_-]?secret|authorization|cookie|set-cookie|password|secret|encryptedKey)(?:[\"'])?\s*[:=]\s*)([\"']?)([^\"',\s}\]]{4,})(\2)",
    re.IGNORECASE,
)
_RE_DATA_IMG = re.compile(r"data:image/[a-zA-Z0-9.+-]+;base64,[A-Za-z0-9+/=]{64,}")
_RE_SECRET_KEY_NAME = re.compile(r"(api[_-]?key|token|secret|password|cookie|authorization|encryptedKey|png_base64|b64_json)", re.IGNORECASE)


def redact_text(value: Any, max_len: int | None = None) -> str:
    text = str(value or "")
    if not text:
        return ""
    text = _RE_BEARER.sub(r"\1[REDACTED]", text)
    text = _RE_SK_KEY.sub("[REDACTED_KEY]", text)
    text = _RE_COOKIE.sub(r"\1=[REDACTED]", text)
    text = _RE_KV_SECRET.sub(r"\1\2[REDACTED]\4", text)
    text = _RE_DATA_IMG.sub("[IMAGE_BASE64_OMITTED]", text)
    if max_len is not None and len(text) > max_len:
        return text[:max_len] + "…(truncated)"
    return text


def sanitize_value(value: Any, depth: int = 0) -> Any:
    if depth > 6:
        return "[max_depth]"
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        return redact_text(value, 160_000)
    if isinstance(value, list):
        return [sanitize_value(item, depth + 1) for item in value[:300]]
    if isinstance(value, dict):
        cleaned: dict[str, Any] = {}
        for k, v in list(value.items())[:120]:
            key = str(k)
            if _RE_SECRET_KEY_NAME.search(key):
                cleaned[key] = "[REDACTED]"
            else:
                cleaned[key] = sanitize_value(v, depth + 1)
        return cleaned
    return redact_text(value, 2000)


def _is_manager(user: dict[str, Any]) -> bool:
    return user.get("role") in {"admin", "manager"}


class FeedbackCreate(BaseModel):
    category: str = Field(default="bug", max_length=32)
    title: str = Field(default="", max_length=160)
    description: str = Field(default="", max_length=6000)
    error_message: str = Field(default="", max_length=2000)
    conversation_id: str = Field(default="", max_length=128)
    route: str = Field(default="", max_length=200)
    client_version: str = Field(default="", max_length=64)
    platform: str = Field(default="", max_length=120)
    client_logs: list[dict[str, Any]] = Field(default_factory=list)
    desktop_diagnostics: dict[str, Any] | None = None
    extra_context: dict[str, Any] | None = None


class FeedbackUpdate(BaseModel):
    status: str | None = Field(default=None, max_length=32)
    admin_note: str | None = Field(default=None, max_length=4000)


def _collect_server_context(db, user: dict[str, Any], conversation_id: str) -> tuple[str, dict[str, Any]]:
    conv_title = ""
    conv_turns: list[dict[str, Any]] = []
    if conversation_id:
        conv = db.one(
            "SELECT id,title,updated_at FROM office_conversations WHERE id=? AND owner_id=?",
            (conversation_id, user["id"]),
        )
        if conv:
            conv_title = str(conv["title"] or "")
            turn_rows = db.all(
                "SELECT id,status,created_at,input,result FROM office_turns WHERE conversation_id=? ORDER BY created_at DESC LIMIT 6",
                (conversation_id,),
            )
            for row in reversed(turn_rows):
                try:
                    inp = unpack(row["input"]) if row["input"] else {}
                except Exception:
                    inp = {}
                try:
                    res = unpack(row["result"]) if row["result"] else {}
                except Exception:
                    res = {}
                conv_turns.append({
                    "turn_id": row["id"],
                    "status": row["status"],
                    "created_at": row["created_at"],
                    "mode": inp.get("mode") or "",
                    "message_preview": redact_text(inp.get("message") or "", 180),
                    "model": res.get("model") or "",
                    "error": redact_text(res.get("error") or "", 600),
                    "steps": [redact_text(s, 200) for s in (res.get("steps") or [])[-20:]] if isinstance(res.get("steps"), list) else [],
                    "has_file": bool(res.get("file_id")),
                })

    run_rows = db.all(
        "SELECT id,kind,engine,status,created_at FROM runs WHERE user_id=? ORDER BY created_at DESC LIMIT 10",
        (user["id"],),
    )
    recent_runs = [
        {
            "id": r["id"],
            "kind": r["kind"],
            "engine": r["engine"],
            "status": r["status"],
            "created_at": r["created_at"],
        }
        for r in run_rows
    ]
    event_rows = db.all(
        "SELECT seq,type,created_at,payload FROM events WHERE actor_id=? ORDER BY seq DESC LIMIT 15",
        (user["id"],),
    )
    recent_events = []
    for ev in event_rows:
        try:
            payload = sanitize_value(unpack(ev["payload"])) if ev["payload"] else {}
        except Exception:
            payload = {}
        recent_events.append({
            "seq": ev["seq"],
            "type": ev["type"],
            "created_at": ev["created_at"],
            "payload": payload,
        })

    return conv_title, {
        "conversation_id": conversation_id if conv_title or conv_turns else "",
        "conversation_title": conv_title,
        "recent_turns": conv_turns,
        "recent_runs": recent_runs,
        "recent_events": recent_events,
    }


def _pack_log_bundle(bundle: dict[str, Any]) -> tuple[str, int]:
    serialized = json.dumps(bundle, ensure_ascii=False, indent=2)
    encoded = serialized.encode("utf-8")
    if len(encoded) <= MAX_LOG_BYTES:
        return serialized, len(encoded)
    # 优先裁减桌面端长文本日志头部，保留尾部最新日志
    trimmed = dict(bundle)
    desktop = dict(trimmed.get("desktop_diagnostics") or {})
    if isinstance(desktop.get("recent_logs"), str) and len(desktop["recent_logs"]) > 60_000:
        desktop["recent_logs"] = "…(较早日志已截断)\n" + desktop["recent_logs"][-60_000:]
        trimmed["desktop_diagnostics"] = desktop
    if isinstance(trimmed.get("client_logs"), list) and len(trimmed["client_logs"]) > 120:
        trimmed["client_logs"] = trimmed["client_logs"][-120:]
    serialized = json.dumps(trimmed, ensure_ascii=False, indent=2)
    encoded = serialized.encode("utf-8")
    if len(encoded) > MAX_LOG_BYTES:
        serialized = encoded[:MAX_LOG_BYTES].decode("utf-8", errors="ignore")
        encoded = serialized.encode("utf-8")
    return serialized, len(encoded)


def _format_log_text(row: dict[str, Any]) -> str:
    lines = [
        "=" * 72,
        f"穹驭办公 问题反馈诊断日志",
        "=" * 72,
        f"反馈编号: {row['id']}",
        f"提交时间: {row['created_at']}",
        f"处理状态: {STATUS_LABELS.get(row['status'], row['status'])}",
        f"问题分类: {CATEGORY_LABELS.get(row['category'], row['category'])}",
        f"对应人: {row['reporter_name']} (ID: {row['reporter_id']})",
        f"部门/角色: {row['reporter_department'] or '未分配'} / {row['reporter_role'] or '未知'}",
        f"邮箱: {row['reporter_email'] or '—'}",
        f"客户端版本: {row['client_version'] or '网页端'} | 服务端版本: {row['server_version'] or '—'}",
        f"运行平台: {row['platform'] or '—'} | 页面路由: {row['route'] or '—'}",
    ]
    if row.get("conversation_id"):
        lines.append(f"关联会话: {row.get('conversation_title') or '未命名会话'} ({row['conversation_id']})")
    lines.extend([
        "-" * 72,
        f"问题标题: {row['title']}",
    ])
    if row.get("error_message"):
        lines.append(f"错误摘要: {row['error_message']}")
    if row.get("description"):
        lines.extend(["", "用户补充说明:", row["description"]])
    if row.get("admin_note"):
        lines.extend(["", f"管理员处理备注 ({row.get('handled_by_name') or row.get('handled_by') or ''}):", row["admin_note"]])

    bundle = {}
    try:
        bundle = json.loads(row.get("log_data") or "{}")
    except Exception:
        bundle = {"raw": row.get("log_data") or ""}

    desktop = bundle.get("desktop_diagnostics") or {}
    if desktop:
        lines.extend(["", "=" * 72, "【一、桌面端诊断信息与运行日志】", "=" * 72])
        lines.append(f"App: {desktop.get('app_version') or '—'} | Electron: {desktop.get('electron_version') or '—'} | Node: {desktop.get('node_version') or '—'}")
        lines.append(f"OS: {desktop.get('os_release') or desktop.get('platform') or '—'} | Origin: {desktop.get('origin') or '—'}")
        if desktop.get("updater_state"):
            lines.append("Updater: " + json.dumps(desktop["updater_state"], ensure_ascii=False))
        if desktop.get("checkpoint_summary"):
            lines.append("Checkpoint: " + json.dumps(desktop["checkpoint_summary"], ensure_ascii=False))
        recent_logs = desktop.get("recent_logs") or ""
        if recent_logs:
            lines.extend(["", "--- 桌面端近期运行日志 ---", str(recent_logs)])

    client_logs = bundle.get("client_logs") or []
    if client_logs:
        lines.extend(["", "=" * 72, f"【二、前端运行日志（共 {len(client_logs)} 条）】", "=" * 72])
        for entry in client_logs:
            ts = entry.get("ts") or ""
            lv = str(entry.get("level") or "info").upper()
            sc = entry.get("scope") or "ui"
            msg = entry.get("message") or ""
            detail = entry.get("detail")
            extra = f" | {json.dumps(detail, ensure_ascii=False)}" if detail else ""
            lines.append(f"[{ts}] [{lv}] [{sc}] {msg}{extra}")

    server_ctx = bundle.get("server_context") or {}
    if server_ctx:
        lines.extend(["", "=" * 72, "【三、服务端关联会话与执行记录】", "=" * 72])
        lines.append(json.dumps(server_ctx, ensure_ascii=False, indent=2))

    extra_ctx = bundle.get("extra_context")
    if extra_ctx:
        lines.extend(["", "=" * 72, "【四、附加现场上下文】", "=" * 72])
        lines.append(json.dumps(extra_ctx, ensure_ascii=False, indent=2))

    return "\n".join(lines) + "\n"


def _row_to_summary(row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "reporter_id": row["reporter_id"],
        "reporter_name": row["reporter_name"],
        "reporter_email": row["reporter_email"],
        "reporter_department": row["reporter_department"],
        "reporter_role": row["reporter_role"],
        "category": row["category"],
        "category_label": CATEGORY_LABELS.get(row["category"], row["category"]),
        "title": row["title"],
        "description": row["description"],
        "error_message": row["error_message"],
        "conversation_id": row["conversation_id"],
        "conversation_title": row["conversation_title"],
        "route": row["route"],
        "client_version": row["client_version"],
        "server_version": row["server_version"],
        "platform": row["platform"],
        "log_summary": row["log_summary"],
        "log_size": row["log_size"],
        "status": row["status"],
        "status_label": STATUS_LABELS.get(row["status"], row["status"]),
        "admin_note": row["admin_note"],
        "handled_by": row["handled_by"],
        "handled_by_name": row["handled_by_name"] if "handled_by_name" in row.keys() else None,
        "handled_at": row["handled_at"],
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


@router.post("")
def submit_feedback(body: FeedbackCreate, user=Depends(user_for)):
    category = body.category.strip().lower() if body.category else "bug"
    if category not in CATEGORIES:
        category = "other"

    error_message = redact_text(body.error_message.strip(), 2000)
    description = redact_text(body.description.strip(), 6000)
    raw_title = redact_text(body.title.strip(), 160)
    if not raw_title:
        if error_message:
            raw_title = error_message.splitlines()[0][:80]
        elif description:
            raw_title = description.splitlines()[0][:80]
        else:
            raw_title = f"{CATEGORY_LABELS.get(category, '问题反馈')}（一键日志上报）"

    conversation_id = (body.conversation_id or "").strip()[:128]
    route = redact_text((body.route or "").strip(), 200)
    client_version = redact_text((body.client_version or "").strip(), 64)
    platform = redact_text((body.platform or "").strip(), 120)

    clean_client_logs = sanitize_value(body.client_logs[-250:]) if body.client_logs else []
    clean_desktop = sanitize_value(body.desktop_diagnostics) if body.desktop_diagnostics else None
    clean_extra = sanitize_value(body.extra_context) if body.extra_context else None

    if not client_version and isinstance(clean_desktop, dict):
        client_version = str(clean_desktop.get("app_version") or "")[:64]
    if not platform and isinstance(clean_desktop, dict):
        platform = str(clean_desktop.get("os_release") or clean_desktop.get("platform") or "")[:120]

    report_id = uid()
    now = instant()

    with database(write=True) as db:
        conv_title, server_ctx = _collect_server_context(db, user, conversation_id)
        bundle = {
            "report_id": report_id,
            "submitted_at": now,
            "reporter": {
                "id": user["id"],
                "name": user["name"],
                "email": user.get("email") or "",
                "department": user.get("department") or "",
                "role": user.get("role") or "",
            },
            "route": route,
            "client_version": client_version,
            "server_version": APP_VERSION,
            "platform": platform,
            "error_message": error_message,
            "desktop_diagnostics": clean_desktop,
            "client_logs": clean_client_logs,
            "server_context": server_ctx,
            "extra_context": clean_extra,
        }
        log_data, log_size = _pack_log_bundle(bundle)
        parts = []
        if isinstance(clean_desktop, dict) and clean_desktop.get("recent_logs"):
            desktop_lines = len(str(clean_desktop["recent_logs"]).splitlines())
            parts.append(f"桌面日志 {desktop_lines} 行")
        if clean_client_logs:
            parts.append(f"前端日志 {len(clean_client_logs)} 条")
        if server_ctx.get("recent_turns"):
            parts.append(f"会话轮次 {len(server_ctx['recent_turns'])} 轮")
        if isinstance(clean_desktop, dict) and clean_desktop.get("checkpoint_summary"):
            parts.append("含断点状态")
        log_summary = " · ".join(parts) if parts else "基础环境信息"

        db.execute(
            """INSERT INTO feedback_reports(
                id, reporter_id, reporter_name, reporter_email, reporter_department, reporter_role,
                category, title, description, error_message, conversation_id, conversation_title,
                route, client_version, server_version, platform, log_summary, log_data, log_size,
                status, admin_note, handled_by, handled_at, created_at, updated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                report_id,
                user["id"],
                user["name"],
                user.get("email") or "",
                user.get("department") or "",
                user.get("role") or "",
                category,
                raw_title,
                description,
                error_message,
                server_ctx.get("conversation_id") or "",
                conv_title,
                route,
                client_version,
                APP_VERSION,
                platform,
                log_summary,
                log_data,
                log_size,
                "open",
                "",
                None,
                None,
                now,
                now,
            ),
        )
        audit(
            db,
            "feedback.submitted",
            user["id"],
            report_id=report_id,
            category=category,
            log_size=log_size,
            reporter_name=user["name"],
        )
        row = db.one(
            """SELECT f.*, h.name AS handled_by_name
               FROM feedback_reports f
               LEFT JOIN users h ON h.id = f.handled_by
               WHERE f.id=?""",
            (report_id,),
        )
    return _row_to_summary(row)


@router.get("")
def list_feedback(
    scope: str = Query(default="auto", max_length=16),
    status: str = Query(default="", max_length=32),
    category: str = Query(default="", max_length=32),
    reporter_id: str = Query(default="", max_length=128),
    q: str = Query(default="", max_length=120),
    offset: int = Query(default=0, ge=0, le=5000),
    limit: int = Query(default=20, ge=1, le=100),
    user=Depends(user_for),
):
    can_manage = _is_manager(user)
    effective_scope = "all" if (can_manage and scope in {"all", "auto"}) else "mine"

    clauses: list[str] = []
    args: list[Any] = []
    if effective_scope == "mine":
        clauses.append("f.reporter_id=?")
        args.append(user["id"])
    elif reporter_id.strip():
        clauses.append("f.reporter_id=?")
        args.append(reporter_id.strip())

    if status.strip() in STATUSES:
        clauses.append("f.status=?")
        args.append(status.strip())
    if category.strip() in CATEGORIES:
        clauses.append("f.category=?")
        args.append(category.strip())
    if q.strip():
        like = f"%{q.strip()}%"
        clauses.append("(f.title LIKE ? OR f.description LIKE ? OR f.error_message LIKE ? OR f.reporter_name LIKE ? OR f.reporter_email LIKE ? OR f.reporter_department LIKE ?)")
        args.extend([like, like, like, like, like, like])

    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""

    with database() as db:
        rows = db.all(
            f"""SELECT f.id, f.reporter_id, f.reporter_name, f.reporter_email, f.reporter_department,
                       f.reporter_role, f.category, f.title, f.description, f.error_message,
                       f.conversation_id, f.conversation_title, f.route, f.client_version,
                       f.server_version, f.platform, f.log_summary, f.log_size, f.status,
                       f.admin_note, f.handled_by, h.name AS handled_by_name, f.handled_at,
                       f.created_at, f.updated_at
                FROM feedback_reports f
                LEFT JOIN users h ON h.id = f.handled_by
                {where}
                ORDER BY f.created_at DESC
                LIMIT ? OFFSET ?""",
            (*args, limit, offset),
        )
        count_where = "WHERE reporter_id=?" if effective_scope == "mine" else ""
        count_args = (user["id"],) if effective_scope == "mine" else ()
        count_rows = db.all(
            f"SELECT status, COUNT(*) AS cnt FROM feedback_reports {count_where} GROUP BY status",
            count_args,
        )
        counts = {"all": 0, "open": 0, "in_progress": 0, "resolved": 0, "closed": 0}
        for cr in count_rows:
            st = cr["status"]
            cnt = int(cr["cnt"] or 0)
            counts["all"] += cnt
            if st in counts:
                counts[st] = cnt

        reporters = []
        if can_manage:
            rep_rows = db.all(
                """SELECT reporter_id, reporter_name, reporter_department, COUNT(*) AS cnt
                   FROM feedback_reports
                   GROUP BY reporter_id, reporter_name, reporter_department
                   ORDER BY cnt DESC, reporter_name ASC"""
            )
            reporters = [
                {
                    "id": r["reporter_id"],
                    "name": r["reporter_name"],
                    "department": r["reporter_department"],
                    "count": int(r["cnt"] or 0),
                }
                for r in rep_rows
            ]

    return {
        "items": [_row_to_summary(row) for row in rows],
        "counts": counts,
        "reporters": reporters,
        "can_manage": can_manage,
        "scope": effective_scope,
    }


def _load_report(db, report_id: str, user: dict[str, Any]):
    row = db.one(
        """SELECT f.*, h.name AS handled_by_name
           FROM feedback_reports f
           LEFT JOIN users h ON h.id = f.handled_by
           WHERE f.id=?""",
        (report_id,),
    )
    if not row:
        fail(404, "反馈记录不存在")
    if not _is_manager(user) and row["reporter_id"] != user["id"]:
        fail(403, "无权查看他人的反馈记录")
    return row


@router.get("/{report_id}")
def get_feedback(report_id: str, user=Depends(user_for)):
    with database() as db:
        row = _load_report(db, report_id, user)
    data = _row_to_summary(row)
    data["log_text"] = _format_log_text(dict(row))
    try:
        data["log_bundle"] = json.loads(row["log_data"] or "{}")
    except Exception:
        data["log_bundle"] = {}
    return data


@router.get("/{report_id}/log")
def download_feedback_log(report_id: str, user=Depends(user_for)):
    with database() as db:
        row = _load_report(db, report_id, user)
    text = _format_log_text(dict(row))
    day = str(row["created_at"] or "")[:10].replace("-", "") or "log"
    safe_name = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", str(row["reporter_name"] or "user"))[:24]
    filename = f"feedback-{day}-{safe_name}-{report_id[:8]}.log"
    ascii_name = f"feedback-{day}-{report_id[:8]}.log"
    return Response(
        content=text.encode("utf-8"),
        media_type="text/plain; charset=utf-8",
        headers={
            "Content-Disposition": f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(filename)}",
            "X-Content-Type-Options": "nosniff",
        },
    )


@router.patch("/{report_id}")
def update_feedback(report_id: str, body: FeedbackUpdate, user=Depends(user_for)):
    if not _is_manager(user):
        fail(403, "仅管理员或主管可更新反馈处理状态")
    with database(write=True) as db:
        row = db.one("SELECT * FROM feedback_reports WHERE id=?", (report_id,))
        if not row:
            fail(404, "反馈记录不存在")
        next_status = row["status"]
        if body.status is not None:
            candidate = body.status.strip().lower()
            if candidate not in STATUSES:
                fail(400, "无效的反馈状态")
            next_status = candidate
        next_note = row["admin_note"] if body.admin_note is None else redact_text(body.admin_note.strip(), 4000)
        now = instant()
        db.execute(
            """UPDATE feedback_reports
               SET status=?, admin_note=?, handled_by=?, handled_at=?, updated_at=?
               WHERE id=?""",
            (next_status, next_note, user["id"], now, now, report_id),
        )
        audit(
            db,
            "feedback.updated",
            user["id"],
            report_id=report_id,
            status=next_status,
            reporter_id=row["reporter_id"],
            reporter_name=row["reporter_name"],
        )
        updated = db.one(
            """SELECT f.*, h.name AS handled_by_name
               FROM feedback_reports f
               LEFT JOIN users h ON h.id = f.handled_by
               WHERE f.id=?""",
            (report_id,),
        )
    return _row_to_summary(updated)
