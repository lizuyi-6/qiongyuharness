import json
import hashlib
import os
import re
import sqlite3
from contextlib import contextmanager
from datetime import UTC, datetime, timedelta
from pathlib import Path
from uuid import uuid4
from zoneinfo import ZoneInfo

ZONE = ZoneInfo(os.getenv("QO_TIMEZONE", "Asia/Shanghai"))
DATA = Path(os.getenv("QO_DATA_DIR", "data"))
DSN = os.getenv("QO_DATABASE_URL", "")
DEMO = os.getenv("QO_DEMO", "0") == "1"

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
 id TEXT PRIMARY KEY, email TEXT, subject TEXT,
 name TEXT NOT NULL, department TEXT NOT NULL, enabled INTEGER NOT NULL DEFAULT 1,
 role TEXT NOT NULL CHECK(role IN ('admin','manager','employee','reviewer')),
 source TEXT NOT NULL DEFAULT 'roster', roster_blocked INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS sessions (
 token TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(id), expires TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS auth_flows (
 state TEXT PRIMARY KEY, verifier TEXT NOT NULL, nonce TEXT NOT NULL,
 desktop TEXT NOT NULL, port INTEGER, browser TEXT NOT NULL, expires TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS desktop_auth_codes (
 code_hash TEXT PRIMARY KEY, challenge TEXT NOT NULL,
 user_id TEXT NOT NULL REFERENCES users(id), expires TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS desktop_auth_codes_expiry ON desktop_auth_codes(expires);
CREATE TABLE IF NOT EXISTS templates (
 id TEXT PRIMARY KEY, owner_id TEXT NOT NULL REFERENCES users(id),
 data TEXT NOT NULL, state TEXT NOT NULL, next_date TEXT,
 created_at TEXT NOT NULL, version INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS assignments (
 id TEXT PRIMARY KEY, template_id TEXT NOT NULL REFERENCES templates(id),
 assignee_id TEXT NOT NULL REFERENCES users(id), owner_id TEXT NOT NULL REFERENCES users(id),
 reviewer_id TEXT REFERENCES users(id), title TEXT NOT NULL, description TEXT NOT NULL,
 work_date TEXT NOT NULL, due_at TEXT NOT NULL, priority TEXT NOT NULL, kind TEXT NOT NULL,
 deliverable TEXT NOT NULL, status TEXT NOT NULL, version INTEGER NOT NULL DEFAULT 1,
 publish_at TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
 UNIQUE(template_id, assignee_id, work_date)
);
CREATE INDEX IF NOT EXISTS assignments_assignee ON assignments(assignee_id, work_date, status);
CREATE INDEX IF NOT EXISTS assignments_owner ON assignments(owner_id, work_date);
CREATE TABLE IF NOT EXISTS runs (
 id TEXT PRIMARY KEY, assignment_id TEXT REFERENCES assignments(id),
 user_id TEXT NOT NULL REFERENCES users(id), kind TEXT NOT NULL,
 engine TEXT NOT NULL, status TEXT NOT NULL, usage TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS artifacts (
 id TEXT PRIMARY KEY, owner_id TEXT NOT NULL REFERENCES users(id),
 assignment_id TEXT REFERENCES assignments(id), template_id TEXT REFERENCES templates(id),
 parent_id TEXT REFERENCES artifacts(id), run_id TEXT REFERENCES runs(id),
 name TEXT NOT NULL, extension TEXT NOT NULL, digest TEXT NOT NULL, size INTEGER NOT NULL,
 data @BLOB@ NOT NULL, label TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS artifacts_task ON artifacts(assignment_id, label);
CREATE TABLE IF NOT EXISTS office_projects (
 id TEXT PRIMARY KEY, owner_id TEXT NOT NULL REFERENCES users(id), name TEXT NOT NULL,
 description TEXT NOT NULL DEFAULT '', brief TEXT NOT NULL DEFAULT '',
 invite_code TEXT NOT NULL UNIQUE, created_at TEXT NOT NULL, updated_at TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS office_project_members (
 project_id TEXT NOT NULL REFERENCES office_projects(id), user_id TEXT NOT NULL REFERENCES users(id),
 joined_at TEXT NOT NULL, PRIMARY KEY(project_id, user_id)
);
CREATE INDEX IF NOT EXISTS office_project_members_user ON office_project_members(user_id);
CREATE TABLE IF NOT EXISTS office_project_context (
 id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES office_projects(id),
 author_id TEXT NOT NULL REFERENCES users(id), title TEXT NOT NULL, content TEXT NOT NULL,
 category TEXT NOT NULL DEFAULT 'doc', digest TEXT NOT NULL DEFAULT '',
 pinned INTEGER NOT NULL DEFAULT 0, enabled INTEGER NOT NULL DEFAULT 1,
 artifact_id TEXT REFERENCES artifacts(id), created_at TEXT NOT NULL, updated_at TEXT NOT NULL DEFAULT '',
 UNIQUE(project_id, artifact_id)
);
CREATE INDEX IF NOT EXISTS office_project_context_project ON office_project_context(project_id, created_at);
CREATE TABLE IF NOT EXISTS office_conversations (
 id TEXT PRIMARY KEY, owner_id TEXT NOT NULL REFERENCES users(id),
 assignment_id TEXT REFERENCES assignments(id), project_id TEXT REFERENCES office_projects(id), title TEXT NOT NULL,
 pinned INTEGER NOT NULL DEFAULT 0, sort_order INTEGER,
 created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS office_conversations_owner ON office_conversations(owner_id, updated_at);
CREATE TABLE IF NOT EXISTS office_turns (
 id TEXT PRIMARY KEY REFERENCES runs(id), conversation_id TEXT NOT NULL REFERENCES office_conversations(id),
 request_id TEXT NOT NULL, input TEXT NOT NULL, result TEXT NOT NULL, status TEXT NOT NULL,
 created_at TEXT NOT NULL, UNIQUE(conversation_id, request_id)
);
CREATE INDEX IF NOT EXISTS office_turns_thread ON office_turns(conversation_id, created_at);
CREATE TABLE IF NOT EXISTS submissions (
 id TEXT PRIMARY KEY, assignment_id TEXT NOT NULL REFERENCES assignments(id),
 author_id TEXT NOT NULL REFERENCES users(id), revision INTEGER NOT NULL, status TEXT NOT NULL,
 note TEXT NOT NULL, link TEXT NOT NULL, review_note TEXT NOT NULL DEFAULT '',
 reviewed_by TEXT REFERENCES users(id), reviewed_at TEXT, created_at TEXT NOT NULL,
 UNIQUE(assignment_id, revision)
);
CREATE TABLE IF NOT EXISTS submission_files (
 submission_id TEXT NOT NULL REFERENCES submissions(id), artifact_id TEXT NOT NULL REFERENCES artifacts(id),
 PRIMARY KEY(submission_id, artifact_id)
);
CREATE TABLE IF NOT EXISTS events (
 seq @SEQUENCE@, assignment_id TEXT REFERENCES assignments(id),
 template_id TEXT REFERENCES templates(id), actor_id TEXT REFERENCES users(id),
 type TEXT NOT NULL, payload TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS events_task ON events(assignment_id, seq);
CREATE TABLE IF NOT EXISTS idempotency (
 actor_id TEXT NOT NULL REFERENCES users(id), action TEXT NOT NULL, key TEXT NOT NULL,
 fingerprint TEXT NOT NULL, result TEXT NOT NULL, created_at TEXT NOT NULL,
 PRIMARY KEY(actor_id, action, key)
);
CREATE TABLE IF NOT EXISTS feedback_reports (
 id TEXT PRIMARY KEY,
 reporter_id TEXT NOT NULL REFERENCES users(id),
 reporter_name TEXT NOT NULL,
 reporter_email TEXT NOT NULL DEFAULT '',
 reporter_department TEXT NOT NULL DEFAULT '',
 reporter_role TEXT NOT NULL DEFAULT '',
 category TEXT NOT NULL DEFAULT 'bug',
 title TEXT NOT NULL,
 description TEXT NOT NULL DEFAULT '',
 error_message TEXT NOT NULL DEFAULT '',
 conversation_id TEXT NOT NULL DEFAULT '',
 conversation_title TEXT NOT NULL DEFAULT '',
 route TEXT NOT NULL DEFAULT '',
 client_version TEXT NOT NULL DEFAULT '',
 server_version TEXT NOT NULL DEFAULT '',
 platform TEXT NOT NULL DEFAULT '',
 log_summary TEXT NOT NULL DEFAULT '',
 log_data TEXT NOT NULL DEFAULT '',
 log_size INTEGER NOT NULL DEFAULT 0,
 status TEXT NOT NULL DEFAULT 'open',
 admin_note TEXT NOT NULL DEFAULT '',
 handled_by TEXT REFERENCES users(id),
 handled_at TEXT,
 created_at TEXT NOT NULL,
 updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS feedback_reports_created ON feedback_reports(created_at);
CREATE INDEX IF NOT EXISTS feedback_reports_reporter ON feedback_reports(reporter_id, created_at);
"""

DEMO_USERS = [
    ("admin", "李祖祎", "企业运营", "admin"),
    ("manager", "周婧", "产品与运营", "manager"),
    ("lin", "林知夏", "产品与运营", "employee"),
    ("chen", "陈以宁", "产品与运营", "employee"),
    ("song", "宋予安", "技术研发", "employee"),
    ("reviewer", "顾言", "产品与运营", "reviewer"),
]


def uid():
    return str(uuid4())


def pack(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def unpack(value):
    return json.loads(value)


def instant(value=None):
    return (value or datetime.now(UTC)).astimezone(UTC).isoformat(timespec="microseconds")


def local_today():
    return datetime.now(ZONE).date()


class Connection:
    def __init__(self, raw):
        self.raw = raw

    def execute(self, query, args=()):
        if DSN:
            query = query.replace("?", "%s")
        return self.raw.execute(query, args)

    def one(self, query, args=()):
        return self.execute(query, args).fetchone()

    def all(self, query, args=()):
        return self.execute(query, args).fetchall()


def connect():
    if DSN:
        import psycopg
        from psycopg.rows import dict_row
        return psycopg.connect(DSN, row_factory=dict_row)
    DATA.mkdir(parents=True, exist_ok=True)
    raw = sqlite3.connect(DATA / "office.db", timeout=15)
    raw.row_factory = sqlite3.Row
    raw.execute("PRAGMA foreign_keys=ON")
    return raw


@contextmanager
def database(write=False):
    raw = connect()
    try:
        conn = Connection(raw)
        if write:
            if DSN:
                conn.execute("SELECT pg_advisory_xact_lock(841684102)")
            else:
                raw.execute("BEGIN IMMEDIATE")
        yield conn
        raw.commit()
    except BaseException:
        raw.rollback()
        raise
    finally:
        raw.close()


def audit(db, kind, actor=None, assignment=None, template=None, **payload):
    db.execute(
        "INSERT INTO events(assignment_id,template_id,actor_id,type,payload,created_at) VALUES(?,?,?,?,?,?)",
        (assignment, template, actor, kind, pack(payload), instant()),
    )


def open_enrollment_enabled():
    value = os.getenv("QO_OPEN_ENROLLMENT", "0")
    if value not in {"0", "1"}:
        raise RuntimeError("QO_OPEN_ENROLLMENT must be exactly '0' or '1'.")
    return value == "1"


def initialize():
    open_enrollment = open_enrollment_enabled()
    with database(write=bool(DSN)) as db:
        schema = SCHEMA.replace("@BLOB@", "BYTEA" if DSN else "BLOB")
        schema = schema.replace("@SEQUENCE@", "BIGSERIAL PRIMARY KEY" if DSN else "INTEGER PRIMARY KEY AUTOINCREMENT")
        if DSN:
            for statement in schema.split(";"):
                if statement.strip():
                    db.execute(statement)
        else:
            db.execute("PRAGMA journal_mode=WAL")
            db.raw.executescript(schema)
    with database(write=True) as db:
        if DSN:
            db.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS enabled INTEGER NOT NULL DEFAULT 1")
            db.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS email TEXT")
            db.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS subject TEXT")
            db.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS source TEXT NOT NULL DEFAULT 'roster'")
            db.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS roster_blocked INTEGER NOT NULL DEFAULT 0")
            db.execute("ALTER TABLE auth_flows ADD COLUMN IF NOT EXISTS port INTEGER")
            db.execute("ALTER TABLE office_conversations ADD COLUMN IF NOT EXISTS project_id TEXT REFERENCES office_projects(id)")
            db.execute("ALTER TABLE office_conversations ADD COLUMN IF NOT EXISTS pinned INTEGER NOT NULL DEFAULT 0")
            db.execute("ALTER TABLE office_conversations ADD COLUMN IF NOT EXISTS sort_order INTEGER")
            db.execute("ALTER TABLE office_projects ADD COLUMN IF NOT EXISTS description TEXT NOT NULL DEFAULT ''")
            db.execute("ALTER TABLE office_projects ADD COLUMN IF NOT EXISTS brief TEXT NOT NULL DEFAULT ''")
            db.execute("ALTER TABLE office_projects ADD COLUMN IF NOT EXISTS updated_at TEXT NOT NULL DEFAULT ''")
            db.execute("ALTER TABLE office_project_context ADD COLUMN IF NOT EXISTS category TEXT NOT NULL DEFAULT 'doc'")
            db.execute("ALTER TABLE office_project_context ADD COLUMN IF NOT EXISTS digest TEXT NOT NULL DEFAULT ''")
            db.execute("ALTER TABLE office_project_context ADD COLUMN IF NOT EXISTS pinned INTEGER NOT NULL DEFAULT 0")
            db.execute("ALTER TABLE office_project_context ADD COLUMN IF NOT EXISTS enabled INTEGER NOT NULL DEFAULT 1")
            db.execute("ALTER TABLE office_project_context ADD COLUMN IF NOT EXISTS updated_at TEXT NOT NULL DEFAULT ''")
        else:
            user_columns = {row["name"] for row in db.all("PRAGMA table_info(users)")}
            if "enabled" not in user_columns:
                db.execute("ALTER TABLE users ADD COLUMN enabled INTEGER NOT NULL DEFAULT 1")
            if "email" not in user_columns:
                db.execute("ALTER TABLE users ADD COLUMN email TEXT")
            if "subject" not in user_columns:
                db.execute("ALTER TABLE users ADD COLUMN subject TEXT")
            if "source" not in user_columns:
                db.execute("ALTER TABLE users ADD COLUMN source TEXT NOT NULL DEFAULT 'roster'")
            if "roster_blocked" not in user_columns:
                db.execute("ALTER TABLE users ADD COLUMN roster_blocked INTEGER NOT NULL DEFAULT 0")
            flow_columns = {row["name"] for row in db.all("PRAGMA table_info(auth_flows)")}
            if "port" not in flow_columns:
                db.execute("ALTER TABLE auth_flows ADD COLUMN port INTEGER")
            conv_columns = {row["name"] for row in db.all("PRAGMA table_info(office_conversations)")}
            if "project_id" not in conv_columns:
                db.execute("ALTER TABLE office_conversations ADD COLUMN project_id TEXT REFERENCES office_projects(id)")
            if "pinned" not in conv_columns:
                db.execute("ALTER TABLE office_conversations ADD COLUMN pinned INTEGER NOT NULL DEFAULT 0")
            if "sort_order" not in conv_columns:
                db.execute("ALTER TABLE office_conversations ADD COLUMN sort_order INTEGER")
            proj_columns = {row["name"] for row in db.all("PRAGMA table_info(office_projects)")}
            if "description" not in proj_columns:
                db.execute("ALTER TABLE office_projects ADD COLUMN description TEXT NOT NULL DEFAULT ''")
            if "brief" not in proj_columns:
                db.execute("ALTER TABLE office_projects ADD COLUMN brief TEXT NOT NULL DEFAULT ''")
            if "updated_at" not in proj_columns:
                db.execute("ALTER TABLE office_projects ADD COLUMN updated_at TEXT NOT NULL DEFAULT ''")
            ctx_columns = {row["name"] for row in db.all("PRAGMA table_info(office_project_context)")}
            if "category" not in ctx_columns:
                db.execute("ALTER TABLE office_project_context ADD COLUMN category TEXT NOT NULL DEFAULT 'doc'")
            if "digest" not in ctx_columns:
                db.execute("ALTER TABLE office_project_context ADD COLUMN digest TEXT NOT NULL DEFAULT ''")
            if "pinned" not in ctx_columns:
                db.execute("ALTER TABLE office_project_context ADD COLUMN pinned INTEGER NOT NULL DEFAULT 0")
            if "enabled" not in ctx_columns:
                db.execute("ALTER TABLE office_project_context ADD COLUMN enabled INTEGER NOT NULL DEFAULT 1")
            if "updated_at" not in ctx_columns:
                db.execute("ALTER TABLE office_project_context ADD COLUMN updated_at TEXT NOT NULL DEFAULT ''")
        db.execute("CREATE UNIQUE INDEX IF NOT EXISTS users_email_unique ON users(email)")
        db.execute("CREATE UNIQUE INDEX IF NOT EXISTS users_subject_unique ON users(subject)")
        users = [(user_id, None, None, name, department, role, True) for user_id, name, department, role in DEMO_USERS] if DEMO else []
        if roster := os.getenv("QO_USERS_FILE"):
            records = json.loads(Path(roster).read_text(encoding="utf-8"))
            emails = set()
            subjects = set()
            for item in records:
                if not isinstance(item, dict) or not all(isinstance(item.get(key), str) and item[key].strip() for key in ("email", "name", "department", "role")):
                    raise ValueError("Invalid QO_USERS_FILE entry")
                email = item["email"].strip().lower()
                if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                    raise ValueError("Invalid QO_USERS_FILE email")
                subject = item.get("sub")
                if subject is not None and (not isinstance(subject, str) or not subject.strip()):
                    raise ValueError("Invalid QO_USERS_FILE subject")
                enabled = item.get("enabled", True)
                if not isinstance(enabled, bool):
                    raise ValueError("Invalid QO_USERS_FILE enabled value")
                subject = subject.strip() if subject is not None else None
                if email in emails or subject is not None and subject in subjects:
                    raise ValueError("Duplicate QO_USERS_FILE email or subject")
                emails.add(email)
                if subject is not None:
                    subjects.add(subject)
                user_id = "u_" + hashlib.sha256(email.encode("utf-8")).hexdigest()[:24]
                users.append((user_id, email, subject, item["name"].strip(), item["department"].strip(), item["role"].strip(), enabled))
        if not DEMO and not any(user[5] == "admin" for user in users):
            raise ValueError("The enterprise roster requires at least one admin.")
        if not DEMO:
            db.execute("UPDATE users SET enabled=0,roster_blocked=0 WHERE source='roster'" if open_enrollment else
                       "UPDATE users SET enabled=0,roster_blocked=0")
        for user_id, email, subject, name, department, role, enabled in users:
            if subject is None:
                db.execute(
                    "INSERT INTO users(id,email,name,department,role,enabled,source,roster_blocked) VALUES(?,?,?,?,?,?,?,?) "
                    "ON CONFLICT(id) DO UPDATE SET email=excluded.email,name=excluded.name,department=excluded.department,role=excluded.role,enabled=excluded.enabled,source='roster',roster_blocked=excluded.roster_blocked",
                    (user_id, email, name, department, role, int(enabled), "roster", int(not enabled)),
                )
            else:
                db.execute(
                    "INSERT INTO users(id,email,subject,name,department,role,enabled,source,roster_blocked) VALUES(?,?,?,?,?,?,?,?,?) "
                    "ON CONFLICT(id) DO UPDATE SET email=excluded.email,subject=excluded.subject,name=excluded.name,department=excluded.department,role=excluded.role,enabled=excluded.enabled,source='roster',roster_blocked=excluded.roster_blocked",
                    (user_id, email, subject, name, department, role, int(enabled), "roster", int(not enabled)),
                )
        if not DEMO:
            db.execute("DELETE FROM sessions WHERE user_id IN (SELECT id FROM users WHERE enabled=0)")
            db.execute("DELETE FROM desktop_auth_codes WHERE user_id IN (SELECT id FROM users WHERE enabled=0)")
        if DEMO and not db.one("SELECT id FROM templates WHERE id=?", ("demo-first",)):
            seed_demo(db)


def seed_demo(db):
    day = local_today()
    rows = [
        ("demo-first", "lin", "制作产品方案汇报 PPT", "结合产品试点目标和现有进展，整理一份可在项目会上汇报的方案。", "pptx", "high", "in_progress", day, 18,
         "交付一份可编辑 PPTX，包含项目背景、方案、实施计划与风险。"),
        ("demo-second", "lin", "整理本周项目例会纪要", "将会议讨论、结论和后续行动整理为内部纪要。", "docx", "normal", "pending", day, 20,
         "交付一份 Word 会议纪要，明确责任人与时间。"),
        ("demo-third", "lin", "更新供应商调研说明", "补充评估依据与下一步验证计划。对外部数据标记出处。", "docx", "urgent", "returned", day - timedelta(days=1), 18,
         "交付 Word 说明，包含结论和待验证事项。"),
        ("demo-fourth", "chen", "整理产品使用反馈", "汇总本周用户反馈并标出优先级。", "docx", "normal", "pending", day, 19, "交付一份可编辑 Word。"),
        ("demo-fifth", "song", "准备版本评审材料", "整理当前版本清单与验收事项。", "pptx", "normal", "pending", day, 21, "交付一份可编辑演示文稿。"),
    ]
    for key, user, title, description, kind, priority, status, work, hour, deliverable in rows:
        now = instant()
        due = instant(datetime(work.year, work.month, work.day, hour, tzinfo=ZONE))
        db.execute("INSERT INTO templates(id,owner_id,data,state,created_at) VALUES(?,?,?,?,?)",
                   (key, "admin", pack({"title": title, "repeat": "once"}), "published", now))
        db.execute(
            """INSERT INTO assignments(id,template_id,assignee_id,owner_id,title,description,work_date,due_at,
               priority,kind,deliverable,status,publish_at,created_at,updated_at)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (key, key, user, "admin", title, description, work.isoformat(), due,
             priority, kind, deliverable, status, now, now, now),
        )
        audit(db, "task.published", "admin", key, key)
        if status == "returned":
            db.execute(
                """INSERT INTO submissions(id,assignment_id,author_id,revision,status,note,link,
                   review_note,reviewed_by,reviewed_at,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                ("demo-review", key, user, 1, "returned", "已提交调研说明初稿。", "", "请补充评估依据与验证计划。", "admin", now, now),
            )
            audit(db, "task.returned", "admin", key, key, note="请补充评估依据与验证计划。")
