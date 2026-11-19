import secrets

from fastapi import APIRouter
from pydantic import BaseModel, Field

from .identity import fail
from .office import clean_text
from .store import audit, database, instant, uid
from .tasks import User

router = APIRouter(prefix="/api/v1/projects")


class ProjectBody(BaseModel):
    name: str = Field(min_length=1, max_length=80)


class JoinBody(BaseModel):
    code: str = Field(min_length=24, max_length=24, pattern=r"^[A-Za-z0-9_-]+$")


class ContextBody(BaseModel):
    title: str = Field(min_length=1, max_length=110)
    content: str = Field(default="", max_length=5000)
    artifact_id: str | None = Field(default=None, max_length=100)


def project_for(db, project_id, user):
    row = db.one("""SELECT p.* FROM office_projects p JOIN office_project_members m ON m.project_id=p.id
                    WHERE p.id=? AND m.user_id=?""", (project_id, user["id"]))
    if not row:
        fail(404, "项目不存在或当前账号无权访问")
    return row


def shared_file_access(db, file_id, user_id):
    return db.one("""SELECT c.id FROM office_project_context c JOIN office_project_members m ON m.project_id=c.project_id
                     WHERE c.artifact_id=? AND m.user_id=? LIMIT 1""", (file_id, user_id)) is not None


def summary(row, user):
    return {"id":row["id"], "name":row["name"], "is_owner":row["owner_id"] == user["id"]}


@router.get("")
def list_projects(user: User):
    with database() as db:
        rows = db.all("""SELECT p.* FROM office_projects p JOIN office_project_members m ON m.project_id=p.id
                         WHERE m.user_id=? ORDER BY m.joined_at DESC LIMIT 60""", (user["id"],))
        return [summary(row,user) for row in rows]


@router.post("")
def create_project(body: ProjectBody, user: User):
    key, now = uid(), instant()
    name = clean_text(body.name, 80)
    with database(write=True) as db:
        db.execute("INSERT INTO office_projects(id,owner_id,name,invite_code,created_at) VALUES(?,?,?,?,?)",
                   (key,user["id"],name,secrets.token_urlsafe(18),now))
        db.execute("INSERT INTO office_project_members(project_id,user_id,joined_at) VALUES(?,?,?)",
                   (key,user["id"],now))
        audit(db,"project.created",user["id"],project_id=key)
        return summary(project_for(db,key,user),user)


@router.post("/join")
def join_project(body: JoinBody, user: User):
    with database(write=True) as db:
        row = db.one("SELECT * FROM office_projects WHERE invite_code=?", (body.code,))
        if not row:
            fail(404,"邀请码无效")
        changed = db.execute("""INSERT INTO office_project_members(project_id,user_id,joined_at) VALUES(?,?,?)
                                ON CONFLICT(project_id,user_id) DO NOTHING""", (row["id"],user["id"],instant())).rowcount
        if changed:
            audit(db,"project.joined",user["id"],project_id=row["id"])
        return summary(row,user)


@router.get("/{key}")
def project_detail(key: str, user: User):
    with database() as db:
        row = project_for(db,key,user)
        members = db.all("""SELECT u.name, u.id FROM office_project_members m JOIN users u ON u.id=m.user_id
                            WHERE m.project_id=? ORDER BY m.joined_at LIMIT 120""",(key,))
        context = db.all("""SELECT c.id,c.title,c.content,c.artifact_id,c.created_at,u.name AS author,a.name AS filename
                            FROM office_project_context c JOIN users u ON u.id=c.author_id
                            LEFT JOIN artifacts a ON a.id=c.artifact_id WHERE c.project_id=?
                            ORDER BY c.created_at DESC LIMIT 50""",(key,))
        return {**summary(row,user), "invite_code":row["invite_code"] if row["owner_id"]==user["id"] else None,
                "members":[dict(m) for m in members], "context":[dict(c) for c in context]}


@router.post("/{key}/invite")
def renew_invite(key: str, user: User):
    with database(write=True) as db:
        row=project_for(db,key,user)
        if row["owner_id"] != user["id"]:
            fail(403,"只有项目创建者可以更新邀请码")
        code=secrets.token_urlsafe(18)
        db.execute("UPDATE office_projects SET invite_code=? WHERE id=?", (code,key))
        return {"invite_code":code}


@router.post("/{key}/context")
def publish_context(key: str, body: ContextBody, user: User):
    title=clean_text(body.title,110)
    content=clean_text(body.content,5000,False)
    if not content and not body.artifact_id:
        fail(422,"请填写项目资料正文或选择文件")
    with database(write=True) as db:
        project_for(db,key,user)
        if body.artifact_id:
            file=db.one("SELECT owner_id,assignment_id,label,extension FROM artifacts WHERE id=?",(body.artifact_id,))
            if not file or file["owner_id"]!=user["id"] or file["assignment_id"] or file["label"]=="material":
                fail(403,"只能共享自己个人工作区里的文件")
            if file["extension"] not in {"docx","pptx","pdf","txt","md"}:
                fail(422,"项目上下文支持 Word、PPT、PDF、TXT 和 Markdown")
            prior=db.one("SELECT id FROM office_project_context WHERE project_id=? AND artifact_id=?",(key,body.artifact_id))
            if prior:
                return {"id":prior["id"]}
        entry=uid()
        db.execute("""INSERT INTO office_project_context(id,project_id,author_id,title,content,artifact_id,created_at)
                      VALUES(?,?,?,?,?,?,?)""",(entry,key,user["id"],title,content,body.artifact_id,instant()))
        audit(db,"project.context_published",user["id"],project_id=key,context_id=entry)
        return {"id":entry}


def project_sources(db, tag, user):
    if not tag:
        return None, [], []
    project=project_for(db,tag,user)
    rows=db.all("""SELECT c.title,c.content,a.name,a.extension,a.data FROM office_project_context c
                   LEFT JOIN artifacts a ON a.id=c.artifact_id WHERE c.project_id=?
                   ORDER BY c.created_at DESC LIMIT 6""",(tag,))
    notes=[]
    files=[]
    remaining=9000
    for row in rows:
        if row["content"] and remaining:
            excerpt=row["content"][:min(2200,remaining)]
            notes.append({"name":"项目："+row["title"],"text":excerpt})
            remaining-=len(excerpt)
        if row["extension"]:
            files.append(("项目："+row["name"],bytes(row["data"]),row["extension"]))
    return project["name"], notes, files
