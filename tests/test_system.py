import json
import os
import tempfile
import unittest
from datetime import UTC, datetime, timedelta
from hashlib import sha256
from io import BytesIO
from uuid import uuid4
from zipfile import ZipFile

TEST_DIR = tempfile.TemporaryDirectory()
os.environ.update(QO_DEMO="1", QO_DATA_DIR=TEST_DIR.name, QO_ENGINE="demo")
if test_dsn := os.getenv("QO_TEST_DATABASE_URL"):
    os.environ["QO_DATABASE_URL"] = test_dsn
else:
    os.environ.pop("QO_DATABASE_URL", None)
os.environ.pop("QO_CLAMD", None)

from docx import Document
from fastapi.testclient import TestClient
from pptx import Presentation

from server.app import app
from server.office import edit_copy, extract, verify, word, Plan, Section
from server.store import local_today

PREFIX = "/api/v1"


class PlatformTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.admin = TestClient(app)
        cls.admin.__enter__()
        cls.accounts = {}
        for name in ("admin", "manager", "lin", "chen", "song", "reviewer"):
            client = cls.admin if name == "admin" else TestClient(app)
            response = client.post(PREFIX + "/auth/demo", json={"user_id": name})
            assert response.status_code == 200, response.text
            cls.accounts[name] = client

    @classmethod
    def tearDownClass(cls):
        cls.admin.__exit__(None, None, None)
        for key, value in cls.accounts.items():
            if key != "admin":
                value.close()

    def key(self):
        return {"Idempotency-Key":str(uuid4())}

    def post(self, actor, url, body, *, key=None, status=200):
        response = self.accounts[actor].post(PREFIX + url, json=body, headers=key or self.key())
        self.assertEqual(response.status_code, status, response.text)
        return response.json()

    def get(self, actor, url, status=200):
        response = self.accounts[actor].get(PREFIX + url)
        self.assertEqual(response.status_code, status, response.text)
        return response.json()

    def create(self, title, *assignees, actor="admin", **options):
        body = {"title":title,"assignee_ids":list(assignees),"work_date":local_today().isoformat(),
                "description":"根据提供的材料整理交付稿。","deliverable":"制作可编辑文件",
                "reviewer_id":None,"publish_mode":"immediate",**options}
        return self.post(actor, "/tasks", body)

    def test_complete_workflow_and_isolation(self):
        args = {"title":"批量派发测试","assignee_ids":["lin","chen"],"work_date":local_today().isoformat(),
                "description":"完成两份内部文书。","deliverable":"可编辑的汇报材料","kind":"pptx",
                "priority":"high","reviewer_id":"reviewer","publish_mode":"draft"}
        created = self.post("admin","/tasks",args)
        material = self.admin.post(PREFIX + f"/tasks/{created['id']}/materials",
                                   files={"file":("项目材料.txt","试点计划与人力安排。".encode(),"text/plain")})
        self.assertEqual(material.status_code,200,material.text)
        material_id = material.json()["id"]
        self.get("lin",f"/files/{material_id}/preview",404)
        key = self.key()
        published = self.post("admin",f"/tasks/{created['id']}/publish",{"version":1},key=key)
        self.assertEqual(published,self.post("admin",f"/tasks/{created['id']}/publish",{"version":1},key=key))
        first, other = published["assignment_ids"]
        self.get("chen",f"/assignments/{first}",404)
        self.get("song",f"/assignments/{first}",404)
        self.get("lin",f"/assignments/{other}",404)
        self.assertIn(first,[t["id"] for t in self.get("lin","/me/today")["items"]])
        self.get("lin",f"/files/{material_id}/preview")
        self.get("song",f"/files/{material_id}/preview",404)
        task = self.get("lin",f"/assignments/{first}")
        start = self.post("lin",f"/assignments/{first}/action/start",{"version":task["version"]})
        self.post("lin",f"/assignments/{first}/action/pause",{"version":task["version"]},status=409)
        self.post("lin",f"/assignments/{first}/action/extension",{"version":start["version"],"note":"等待资料确认"})
        results = self.post("lin",f"/assignments/{first}/sessions",
                            {"kind":"pptx","title":"测试演示","instructions":"按阶段组织内容","pages":12,"material_ids":[material_id]})
        self.assertEqual(len(results["plan"]["sections"]),11)
        ppt = self.post("lin","/agent/render",{"kind":"pptx","plan":results["plan"],"run_id":results["run_id"],"assignment_id":first})
        path = PREFIX + f"/files/{ppt['id']}/download"
        response = self.accounts["lin"].get(path)
        self.assertEqual(response.status_code,200,response.text[:150] if response.status_code!=200 else "")
        prs = Presentation(BytesIO(response.content))
        self.assertEqual(len(prs.slides),12)
        self.assertEqual(sha256(response.content).hexdigest(),ppt["digest"])
        self.assertEqual(self.accounts["reviewer"].get(path).status_code,404)
        self.assertEqual(self.accounts["admin"].get(path).status_code,404)
        key = self.key()
        body = {"version":start["version"],"note":"首轮提交","artifact_ids":[ppt["id"]]}
        sub = self.post("lin",f"/assignments/{first}/submissions",body,key=key)
        self.assertEqual(sub,self.post("lin",f"/assignments/{first}/submissions",body,key=key))
        self.assertEqual(self.accounts["reviewer"].get(path).status_code,200)
        detail = self.get("reviewer",f"/assignments/{first}")
        self.post("chen",f"/submissions/{sub['id']}/review",{"version":detail["version"],"decision":"approve"},status=404)
        self.post("reviewer",f"/submissions/{sub['id']}/review",{"version":detail["version"],"decision":"return"},status=422)
        self.post("reviewer",f"/submissions/{sub['id']}/review",{"version":detail["version"],"decision":"return","note":"请核对项目名称"})
        returned = self.get("lin",f"/assignments/{first}")
        self.assertEqual(returned["status"],"returned")
        edited = self.post("lin",f"/files/{ppt['id']}/edit",{"old":"测试演示","new":"项目汇报正式版"})
        self.assertEqual(edited["parent_id"],ppt["id"])
        self.assertEqual(self.accounts["lin"].get(path).content,response.content)
        self.assertIn("项目汇报正式版",self.get("lin",f"/files/{edited['id']}/preview")["text"])
        sub2 = self.post("lin",f"/assignments/{first}/submissions",{"version":returned["version"],"artifact_ids":[edited["id"]]})
        reopened = self.get("admin",f"/assignments/{first}")
        self.post("admin",f"/submissions/{sub2['id']}/review",{"version":reopened["version"],"decision":"approve"})
        done = self.get("lin",f"/assignments/{first}")
        self.assertEqual(done["status"],"approved")
        self.assertEqual([v["status"] for v in done["submissions"]],["approved","returned"])
        self.assertEqual(self.get("chen",f"/assignments/{other}")["status"],"pending")

    def test_schedule_visibility_timezone_and_scope(self):
        now = datetime.now(UTC)
        key = self.create("未来发布", "lin", publish_at=(now+timedelta(minutes=10)).isoformat())["assignment_ids"][0]
        self.get("lin",f"/assignments/{key}",404)
        self.assertEqual(self.get("admin",f"/assignments/{key}")["status"],"scheduled")
        overdue = self.create("昨日任务","lin",work_date=(local_today()-timedelta(days=1)).isoformat(),
                              due_at=(now-timedelta(hours=25)).isoformat())["assignment_ids"][0]
        self.assertIn(overdue,[a["id"] for a in self.get("lin","/me/today")["items"] if a["overdue"]])
        self.assertEqual(self.get("lin","/me/today")["timezone"],"Asia/Shanghai")
        data={"title":"跨部门派发","work_date":local_today().isoformat(),"assignee_ids":["song"]}
        self.post("manager","/tasks",data,status=403)
        recurring = self.create("定时工作日","lin","chen",repeat="weekdays",
                                end_date=(local_today()+timedelta(days=2)).isoformat(),
                                include_dates=[local_today().isoformat()])
        tasks=self.get("admin","/admin/tasks?limit=200")
        self.assertEqual(len([a for a in tasks if a["template_id"]==recurring["id"]]),2)
        self.assertEqual(len([a for a in self.get("admin","/admin/tasks?limit=200") if a["template_id"]==recurring["id"]]),2)

    def test_docx_and_draft_versions(self):
        plan=Plan(title="文件标题",sections=[Section(heading="公司报告",points=["文本保持可编辑"])])
        data=word(plan)
        self.assertEqual(verify("测试.docx",data),("测试.docx","docx"))
        self.assertIn("文本保持可编辑","\n".join(p.text for p in Document(BytesIO(data)).paragraphs))
        doc=Document(); p=doc.add_paragraph()
        p.add_run("别覆盖").bold=True
        p.add_run("原稿").italic=True
        buf=BytesIO(); doc.save(buf)
        new,count=edit_copy(buf.getvalue(),"docx","覆盖原","改写旧")
        self.assertEqual(count,1)
        self.assertIn("别改写旧稿",extract(new,"docx"))
        self.assertIn("别覆盖原稿",extract(buf.getvalue(),"docx"))

    def test_idempotency_reassignment_and_origin(self):
        first = self.create("转派测试","lin")["assignment_ids"][0]
        initial = self.get("lin","/me/tasks")
        old = self.get("lin",f"/assignments/{first}")
        response = self.admin.patch(PREFIX+f"/assignments/{first}",json={"version":old["version"],"assignee_id":"chen"})
        self.assertEqual(response.status_code,200,response.text)
        self.get("lin",f"/assignments/{first}",404)
        delta = self.get("lin",f"/me/tasks?cursor={initial['cursor']}")
        self.assertIn({"id":first,"removed":True},delta["items"])
        self.get("chen",f"/assignments/{first}")
        response=self.admin.post(PREFIX+"/auth/logout",headers={"Origin":"https://untrusted.invalid"},json={})
        self.assertEqual(response.status_code,403)
        self.get("admin","/admin/audit")
        self.get("lin","/admin/audit",403)

    def test_oidc_desktop_challenge_and_single_use(self):
        import json
        import secrets
        from unittest.mock import patch
        from urllib.parse import parse_qs, urlsplit
        import httpx
        import jwt
        from cryptography.hazmat.primitives.asymmetric import rsa
        from server import identity, store

        issuer="https://identity.test"
        application="office-test"
        signing_key=rsa.generate_private_key(public_exponent=65537,key_size=2048)
        jwk=json.loads(jwt.algorithms.RSAAlgorithm.to_jwk(signing_key.public_key()))
        jwk.update({"kid":"mock-key","use":"sig","alg":"RS256"})
        action={}
        attempts=[]
        claims_by_code={}
        real_client=httpx.AsyncClient

        def endpoint(request):
            if request.url.path=="/.well-known/openid-configuration":
                return httpx.Response(200,json={"issuer":issuer,"authorization_endpoint":issuer+"/authorize",
                    "token_endpoint":issuer+"/token","jwks_uri":issuer+"/keys"})
            if request.url.path=="/keys":
                return httpx.Response(200,json={"keys":[jwk]})
            if request.url.path=="/token":
                form=parse_qs(request.content.decode())
                attempts.append(form["code_verifier"][0])
                self.assertEqual(identity.challenge(form["code_verifier"][0]), action["code_challenge"][0])
                now=datetime.now(UTC)
                claims={"iss":issuer,"sub":"lin-subject","aud":application,"iat":now,
                        "exp":now+timedelta(minutes=4),"nonce":action["nonce"][0],
                        **claims_by_code[form["code"][0]]}
                return httpx.Response(200,json={"id_token":jwt.encode(claims,signing_key,algorithm="RS256",headers={"kid":"mock-key"})})
            return httpx.Response(404)

        verifier="b"*64
        challenge=identity.challenge(verifier)
        with store.database(write=True) as db:
            db.execute("UPDATE users SET email=?,subject=NULL WHERE id=?",("lin@example.com","lin"))
        with (patch.object(identity,"ISSUER",issuer),patch.object(identity,"CLIENT",application),
              patch.object(identity,"CLIENT_SECRET",""),
              patch.object(identity.httpx,"AsyncClient",side_effect=lambda **kw:real_client(transport=httpx.MockTransport(endpoint),**kw))):
            browser=TestClient(app,follow_redirects=False)
            frontend=TestClient(app)
            for invalid_port in ("abc", "1023", "70000"):
                self.assertEqual(browser.get("/auth/start",params={"desktop":challenge,"port":invalid_port}).status_code,400)
            self.assertEqual(browser.get("/auth/start",params={"desktop":challenge}).status_code,400)
            response=browser.get("/auth/start",params={"desktop":challenge,"port":"43127"})
            self.assertEqual(response.status_code,303,response.text)
            action.update(parse_qs(urlsplit(response.headers["location"]).query))
            self.assertEqual(action["code_challenge_method"],["S256"])
            self.assertEqual(action["scope"],["openid profile email"])
            claims_by_code["mock-code"]={"email":"lin@example.com","email_verified":True}
            params={"state":action["state"][0],"code":"mock-code"}
            self.assertEqual(frontend.get("/auth/callback",params=params).status_code,400)
            callback=browser.get("/auth/callback",params=params)
            self.assertEqual(callback.status_code,303)
            self.assertEqual(urlsplit(callback.headers["location"]).netloc,"127.0.0.1:43127")
            self.assertEqual(urlsplit(callback.headers["location"]).path,"/callback")
            auth_code=parse_qs(urlsplit(callback.headers["location"]).query)["code"][0]
            self.assertGreaterEqual(len(auth_code),43)
            self.assertEqual(browser.get("/auth/callback",params=params).status_code,400)
            self.assertEqual(frontend.post(PREFIX+"/auth/desktop/claim",json={"code":auth_code,"verifier":"a"*64}).status_code,400)
            claim=frontend.post(PREFIX+"/auth/desktop/claim",json={"code":auth_code,"verifier":verifier})
            self.assertEqual(claim.status_code,200,claim.text)
            self.assertEqual(frontend.post(PREFIX+"/auth/desktop/claim",json={"code":auth_code,"verifier":verifier}).status_code,404)
            with store.database() as db:
                self.assertEqual(db.one("SELECT subject FROM users WHERE id=?",("lin",))["subject"],"lin-subject")
            expired_code=secrets.token_urlsafe(32)
            with store.database(write=True) as db:
                db.execute("INSERT INTO desktop_auth_codes(code_hash,challenge,user_id,expires) VALUES(?,?,?,?)",
                           (identity.digest(expired_code),challenge,"lin",store.instant(datetime.now(UTC)-timedelta(minutes=1))))
            self.assertEqual(frontend.post(PREFIX+"/auth/desktop/claim",json={"code":expired_code,"verifier":verifier}).status_code,404)
            frontend.cookies.set(identity.COOKIE,claim.json()["session"])
            self.assertEqual(frontend.get(PREFIX+"/me").json()["id"],"lin")
            claims_by_code["subject-login"]={"email_verified":False}
            response=browser.get("/auth/start")
            action.update(parse_qs(urlsplit(response.headers["location"]).query))
            response=browser.get("/auth/callback",params={"state":action["state"][0],"code":"subject-login"})
            self.assertEqual(response.status_code,303,response.text)
            self.assertEqual(response.headers["location"],"/")
            self.assertEqual(len(attempts),2)
            frontend.close(); browser.close()
        with store.database(write=True) as db:
            db.execute("UPDATE users SET email=NULL,subject=NULL WHERE id=?",("lin",))

    def test_conversation_with_task_sources_and_revised_output(self):
        assignment = self.create("试点汇报","lin",kind="pptx")["assignment_ids"][0]
        thread = str(uuid4())
        self.post("chen","/conversations",{"id":str(uuid4()),"assignment_id":assignment},status=404)
        self.post("lin","/conversations",{"id":thread,"assignment_id":assignment})
        self.get("chen",f"/conversations/{thread}",404)
        upload = self.accounts["lin"].post(PREFIX + "/files",files={"file":(
            "调研记录.txt","安排下周三复核渠道反馈，负责人待定。".encode(),"text/plain")},data={"assignment_id":assignment})
        self.assertEqual(upload.status_code,200,upload.text)
        material_id = upload.json()["id"]

        def turn(message, mode="auto", same_id=None):
            body = {"id":same_id or str(uuid4()),"message":message,"mode":mode,"material_ids":[material_id]}
            res = self.accounts["lin"].post(PREFIX + f"/conversations/{thread}/turns",json=body)
            self.assertEqual(res.status_code,200,res.text)
            events = [json.loads(line) for line in res.text.splitlines()]
            self.assertEqual(events[-1]["type"],"complete",events)
            return events[-1]["turn"]

        initial = turn("只总结资料，给出待办清单","reply")
        self.assertIn("复核渠道反馈",initial["reply"])
        self.assertNotIn("file",initial)
        message = "请生成一份可编辑的 5 页 PPT"
        request_id = str(uuid4())
        exported = turn(message,same_id=request_id)
        self.assertEqual(turn(message,same_id=request_id)["file"]["id"],exported["file"]["id"])
        first_file = exported["file"]
        self.assertIn("已导出 PowerPoint",exported["steps"])
        raw = self.accounts["lin"].get(PREFIX + f"/files/{first_file['id']}/download").content
        self.assertEqual(len(Presentation(BytesIO(raw)).slides),5)
        self.assertIn("复核渠道反馈",extract(raw,"pptx"))
        revised = turn("补上负责人核对事项，重新导出")
        self.assertEqual(revised["file"]["parent_id"],first_file["id"])
        result = self.get("lin",f"/conversations/{thread}")
        self.assertEqual(len(result["turns"]),3)
        self.assertEqual(result["turns"][-1]["file"]["assignment_id"],assignment)
        self.assertTrue(any(item["id"]==thread for item in self.get("lin","/conversations")))
        self.assertFalse(any(item["id"]==thread for item in self.get("chen","/conversations")))

    def test_shared_project_context_and_slide_scoped_revision(self):
        group=self.post("lin","/projects",{"name":"联调项目"})
        project_id=group["id"]
        code=self.get("lin",f"/projects/{project_id}")["invite_code"]
        self.get("chen",f"/projects/{project_id}",404)
        self.post("chen","/conversations",{"id":str(uuid4()),"project_id":project_id},status=404)
        self.post("chen","/projects/join",{"code":code})
        self.assertIsNone(self.get("chen",f"/projects/{project_id}")["invite_code"])
        self.post("lin",f"/projects/{project_id}/context",{"title":"评审决策","content":"验收改到周四下午。"})
        attachment=self.accounts["lin"].post(PREFIX+"/files",files={"file":("检查项.txt","需核对渠道数据。".encode(),"text/plain")}).json()
        self.post("lin",f"/projects/{project_id}/context",{"title":"共同核对清单","artifact_id":attachment["id"]})
        self.get("song",f"/files/{attachment['id']}/preview",404)
        self.assertIn("渠道数据",self.get("chen",f"/files/{attachment['id']}/preview")["text"])
        secret=str(uuid4())
        self.post("lin","/conversations",{"id":secret,"project_id":project_id})
        self.get("chen",f"/conversations/{secret}",404)
        thread=str(uuid4())
        self.post("chen","/conversations",{"id":thread,"project_id":project_id})

        def turn(message,mode="auto",selection=None):
            payload={"id":str(uuid4()),"message":message,"mode":mode}
            if selection:
                payload["selection"]=selection
            response=self.accounts["chen"].post(PREFIX+f"/conversations/{thread}/turns",json=payload)
            self.assertEqual(response.status_code,200,response.text)
            lines=[json.loads(line) for line in response.text.splitlines()]
            self.assertEqual(lines[-1]["type"],"complete",lines)
            return lines[-1]["turn"]

        answer=turn("列出项目安排","reply")
        self.assertIn("周四下午",answer["reply"])
        self.assertIn("渠道数据",answer["reply"])
        original=turn("生成 5 页 PPT")["file"]
        file_id=original["id"]
        self.assertEqual(len(self.get("chen",f"/files/{file_id}/slides")["pages"]),5)
        self.get("lin",f"/files/{file_id}/slides/2",404)
        scene=self.get("chen",f"/files/{file_id}/slides/2")
        self.assertGreater(len(scene["objects"]),2)
        target=next(obj for obj in scene["objects"] if obj.get("editable"))
        revised=turn("把所选文本改为「试点执行口径」",selection={"file_id":file_id,"slide":2,"shape_id":target["id"]})["file"]
        self.assertEqual(revised["parent_id"],file_id)
        before=self.accounts["chen"].get(PREFIX+f"/files/{file_id}/download").content
        after=self.accounts["chen"].get(PREFIX+f"/files/{revised['id']}/download").content
        with ZipFile(BytesIO(before)) as old,ZipFile(BytesIO(after)) as new:
            self.assertEqual(old.namelist(),new.namelist())
            changed=[name for name in old.namelist() if old.read(name)!=new.read(name)]
        self.assertEqual(changed,["ppt/slides/slide2.xml"])
        before_blocks={o["id"]:o["text"] for o in scene["objects"] if o.get("editable")}
        after_blocks={o["id"]:o["text"] for o in self.get("chen",f"/files/{revised['id']}/slides/2")["objects"] if o.get("editable")}
        self.assertEqual([i for i in before_blocks if before_blocks[i]!=after_blocks[i]],[target["id"]])
        self.assertEqual(after_blocks[target["id"]],"试点执行口径")
