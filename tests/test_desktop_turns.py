import json
import os
import tempfile
import unittest
from uuid import uuid4
from io import BytesIO

from fastapi.testclient import TestClient
from pptx import Presentation

TEST_DATA = tempfile.TemporaryDirectory()
os.environ.update(QO_DEMO="1", QO_DATA_DIR=TEST_DATA.name, QO_ENGINE="demo")
os.environ.pop("QO_DATABASE_URL", None)
os.environ.pop("QO_CLAMD", None)


class DesktopTurnTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # test_system.py initializes the shared demo database before this class runs.
        from server import gateway
        from server.app import app

        cls.gateway = gateway
        cls.original_mode = gateway.MODE
        gateway.MODE = "desktop"
        cls.clients = {}
        for name in ("admin", "lin", "chen"):
            client = TestClient(app)
            client.__enter__()
            response = client.post("/api/v1/auth/demo", json={"user_id": name})
            if response.status_code != 200:
                raise AssertionError(response.text)
            cls.clients[name] = client

    @classmethod
    def tearDownClass(cls):
        for client in cls.clients.values():
            client.__exit__(None, None, None)
        cls.gateway.MODE = cls.original_mode

    def conversation(self, actor="lin"):
        key = str(uuid4())
        response = self.clients[actor].post("/api/v1/conversations", json={"id": key})
        self.assertEqual(response.status_code, 200, response.text)
        return key

    def prepare(self, key, actor="lin", message="请生成 3 页 PPT"):
        response = self.clients[actor].post(f"/api/v1/conversations/{key}/turns/prepare", json={
            "id": str(uuid4()), "message": message, "mode": "pptx"
        })
        self.assertEqual(response.status_code, 200, response.text)
        prepared = response.json()
        self.assertEqual(prepared["kind"], "conversation")
        self.assertIn("system", prepared)
        self.assertIn("prompt", prepared)
        return prepared

    def complete(self, key, prepared, actor="lin", text=None):
        value = text or json.dumps({
            "action": "export", "reply": "已整理演示稿", "kind": "pptx",
            "plan": {"title": "桌面执行验证", "subtitle": "", "sections": [
                {"heading": "目标", "points": ["核对接入方案"]},
                {"heading": "安排", "points": ["补充验证记录"]},
            ]},
        }, ensure_ascii=False)
        return self.clients[actor].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/complete",
            json={"text": value, "usage": {"prompt_tokens": 13, "completion_tokens": 9, "total_tokens": 22},
                  "model": "desktop-test-model"},
        )

    def test_prepare_complete_exports_and_records_model_usage(self):
        key = self.conversation()
        prepared = self.prepare(key)
        response = self.complete(key, prepared)
        self.assertEqual(response.status_code, 200, response.text)
        turn = response.json()
        self.assertEqual(turn["status"], "completed")
        self.assertTrue(turn["file"]["name"].endswith(".pptx"))
        duplicate = self.complete(key, prepared)
        self.assertEqual(duplicate.status_code, 409, duplicate.text)
        self.assertTrue(turn["file"]["id"])
        download = self.clients["lin"].get(f"/api/v1/files/{turn['file']['id']}/download")
        self.assertEqual(download.status_code, 200, download.content[:80])
        self.assertEqual(len(Presentation(BytesIO(download.content)).slides), 3)
        from server.store import database, unpack
        with database() as db:
            usage = unpack(db.one("SELECT usage FROM runs WHERE id=?", (prepared["run_id"],))["usage"])
        self.assertEqual(usage["model"], "desktop-test-model")
        self.assertEqual(usage["total_tokens"], 22)

    def test_invalid_model_json_fails_run_and_second_completion_conflicts(self):
        key = self.conversation()
        prepared = self.prepare(key)
        rejected = self.complete(key, prepared, text="not a JSON decision")
        self.assertEqual(rejected.status_code, 502, rejected.text)
        again = self.complete(key, prepared)
        self.assertEqual(again.status_code, 409, again.text)
        turn = self.clients["lin"].get(f"/api/v1/conversations/{key}").json()["turns"][-1]
        self.assertEqual(turn["status"], "failed")
        self.assertIn("error", turn)

    def test_slide_selection_is_revised_inside_the_authorized_ppt(self):
        key = self.conversation()
        first = self.prepare(key)
        original_response = self.complete(key, first)
        self.assertEqual(original_response.status_code, 200, original_response.text)
        original = original_response.json()["file"]
        page = self.clients["lin"].get(f"/api/v1/files/{original['id']}/slides/2")
        self.assertEqual(page.status_code, 200, page.text)
        target = next(item for item in page.json()["objects"] if item.get("editable"))
        request = {
            "id": str(uuid4()), "message": "把选中文字改成修订事项", "mode": "auto",
            "selection": {"file_id": original["id"], "slide": 2, "shape_id": target["id"]},
        }
        prepared_response = self.clients["lin"].post(f"/api/v1/conversations/{key}/turns/prepare", json=request)
        self.assertEqual(prepared_response.status_code, 200, prepared_response.text)
        prepared = prepared_response.json()
        self.assertEqual(prepared["kind"], "slide")
        change = json.dumps({"reply": "已按选区修订", "changes": [
            {"shape_id": target["id"], "text": "修订事项"},
        ]}, ensure_ascii=False)
        completed = self.clients["lin"].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/complete",
            json={"text": change, "usage": {"total_tokens": 4}, "model": "slide-test-model"},
        )
        self.assertEqual(completed.status_code, 200, completed.text)
        revised = completed.json()["file"]
        self.assertEqual(revised["parent_id"], original["id"])
        scene = self.clients["lin"].get(f"/api/v1/files/{revised['id']}/slides/2")
        self.assertEqual(scene.status_code, 200, scene.text)
        blocks = [item["text"] for item in scene.json()["objects"] if item.get("editable")]
        self.assertIn("修订事项", blocks)

    def test_other_account_cannot_complete_and_web_execution_is_rejected(self):
        key = self.conversation()
        prepared = self.prepare(key)
        other = self.complete(key, prepared, actor="chen")
        self.assertEqual(other.status_code, 404, other.text)
        direct = self.clients["lin"].post(f"/api/v1/conversations/{key}/turns", json={
            "id": str(uuid4()), "message": "请回复", "mode": "reply"
        })
        self.assertEqual(direct.status_code, 409, direct.text)
        owner = self.complete(key, prepared)
        self.assertEqual(owner.status_code, 200, owner.text)

    def test_fail_endpoint_marks_running_turn_failed(self):
        key = self.conversation()
        request = {"id": str(uuid4()), "message": "请生成 3 页 PPT", "mode": "pptx"}
        response = self.clients["lin"].post(f"/api/v1/conversations/{key}/turns/prepare", json=request)
        self.assertEqual(response.status_code, 200, response.text)
        prepared = response.json()
        same_request = self.clients["lin"].post(f"/api/v1/conversations/{key}/turns/prepare", json=request)
        self.assertEqual(same_request.status_code, 409, same_request.text)
        locked = self.clients["lin"].post(f"/api/v1/conversations/{key}/turns/prepare", json={
            "id": str(uuid4()), "message": "另一轮请求", "mode": "reply"
        })
        self.assertEqual(locked.status_code, 409, locked.text)
        failed = self.clients["lin"].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/fail", json={"message": "桌面执行失败"}
        )
        self.assertEqual(failed.status_code, 200, failed.text)
        self.assertEqual(failed.json(), {"ok": True})
        duplicate = self.complete(key, prepared)
        self.assertEqual(duplicate.status_code, 409, duplicate.text)

    def test_agent_plan_prepare_and_complete(self):
        prepared_response = self.clients["lin"].post("/api/v1/agent/plan/prepare", json={
            "kind": "pptx", "title": "桌面计划", "instructions": "整理两项安排", "pages": 3,
        })
        self.assertEqual(prepared_response.status_code, 200, prepared_response.text)
        prepared = prepared_response.json()
        payload = {"run_id": prepared["run_id"], "text": json.dumps({
            "title": "桌面计划", "subtitle": "", "sections": [
                {"heading": "第一项", "points": ["核对材料"]},
                {"heading": "第二项", "points": ["记录结论"]},
            ]}, ensure_ascii=False), "usage": {"total_tokens": 7}, "model": "plan-test-model"}
        completed = self.clients["lin"].post("/api/v1/agent/plan/complete", json=payload)
        self.assertEqual(completed.status_code, 200, completed.text)
        self.assertEqual(len(completed.json()["plan"]["sections"]), 2)
        self.assertEqual(completed.json()["engine"], "desktop")
        from server.store import database, unpack
        with database() as db:
            usage = unpack(db.one("SELECT usage FROM runs WHERE id=?", (prepared["run_id"],))["usage"])
        self.assertEqual(usage["model"], "plan-test-model")

    def test_assignment_session_prepare_and_complete(self):
        from server.store import local_today
        created = self.clients["admin"].post("/api/v1/tasks", json={
            "title": "桌面计划任务", "assignee_ids": ["lin"],
            "work_date": local_today().isoformat(), "publish_mode": "immediate",
        }, headers={"Idempotency-Key": str(uuid4())})
        self.assertEqual(created.status_code, 200, created.text)
        assignment = created.json()["assignment_ids"][0]
        prepared_response = self.clients["lin"].post(f"/api/v1/assignments/{assignment}/sessions/prepare", json={
            "kind": "docx", "title": "任务文稿", "instructions": "整理任务安排",
        })
        self.assertEqual(prepared_response.status_code, 200, prepared_response.text)
        prepared = prepared_response.json()
        text = json.dumps({"title": "任务文稿", "subtitle": "", "sections": [
            {"heading": "安排", "points": ["完成资料核对"]},
        ]}, ensure_ascii=False)
        completed = self.clients["lin"].post(f"/api/v1/assignments/{assignment}/sessions/complete", json={
            "run_id": prepared["run_id"], "text": text, "usage": {"total_tokens": 5},
            "model": "assignment-test-model",
        })
        self.assertEqual(completed.status_code, 200, completed.text)
        self.assertEqual(completed.json()["engine"], "desktop")

    def test_engine_advertises_client_execution(self):
        response = self.clients["lin"].get("/api/v1/engine")
        self.assertEqual(response.status_code, 200, response.text)
        self.assertIs(response.json()["client_execution"], True)
        previous = (self.gateway.URL, self.gateway.KEY)
        try:
            self.gateway.URL = ""
            self.gateway.KEY = ""
            self.gateway.validate()
        finally:
            self.gateway.URL, self.gateway.KEY = previous


if __name__ == "__main__":
    unittest.main()
