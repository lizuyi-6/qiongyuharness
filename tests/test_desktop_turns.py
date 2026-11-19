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

    def complete(self, key, prepared, actor="lin", text=None, usage=None):
        value = text or json.dumps({
            "action": "export", "reply": "已整理演示稿", "kind": "pptx",
            "plan": {"title": "桌面执行验证", "subtitle": "", "sections": [
                {"heading": "目标", "points": ["核对接入方案"]},
                {"heading": "安排", "points": ["补充验证记录"]},
            ]},
        }, ensure_ascii=False)
        return self.clients[actor].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/complete",
            json={"text": value, "usage": usage or {"prompt_tokens": 13, "completion_tokens": 9, "total_tokens": 22},
                  "model": "desktop-test-model"},
        )

    def test_prepare_complete_exports_and_records_model_usage(self):
        key = self.conversation()
        prepared = self.prepare(key)
        telemetry = {
            "prompt_tokens": 13, "completion_tokens": 9, "total_tokens": 22,
            "max_tokens_clamped": [{"requested": 5000, "applied": 3500}],
            "finish_reasons": ["length"], "truncated": True,
        }
        response = self.complete(key, prepared, usage=telemetry)
        self.assertEqual(response.status_code, 200, response.text)
        turn = response.json()
        self.assertEqual(turn["status"], "completed")
        self.assertEqual(turn["output_telemetry"], {key: telemetry[key] for key in
                         ("max_tokens_clamped", "finish_reasons", "truncated")})
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
            audit = unpack(db.one("SELECT payload FROM events WHERE type='agent.completed' AND payload LIKE ?",
                                  (f'%{prepared["run_id"]}%',))["payload"])
        self.assertEqual(usage["model"], "desktop-test-model")
        self.assertEqual(usage["total_tokens"], 22)
        self.assertEqual(usage["max_tokens_clamped"], [{"requested": 5000, "applied": 3500}])
        self.assertEqual(usage["finish_reasons"], ["length"])
        self.assertIs(usage["truncated"], True)
        self.assertEqual(audit["usage"]["max_tokens_clamped"], [{"requested": 5000, "applied": 3500}])
        self.assertEqual(audit["usage"]["finish_reasons"], ["length"])
        self.assertIs(audit["usage"]["truncated"], True)

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

    def create_assignment(self):
        from server.store import local_today
        created = self.clients["admin"].post("/api/v1/tasks", json={
            "title": "文稿失败路径验证", "assignee_ids": ["lin"],
            "work_date": local_today().isoformat(), "publish_mode": "immediate",
        }, headers={"Idempotency-Key": str(uuid4())})
        self.assertEqual(created.status_code, 200, created.text)
        return created.json()["assignment_ids"][0]

    def prepare_agent_plan(self, assignment=None):
        path = f"/api/v1/assignments/{assignment}/sessions/prepare" if assignment else "/api/v1/agent/plan/prepare"
        response = self.clients["lin"].post(path, json={
            "kind": "docx", "title": "失败状态验证", "instructions": "验证桌面执行失败记录",
        })
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()

    def test_agent_plan_fail_marks_run_failed_and_blocks_complete(self):
        prepared = self.prepare_agent_plan()
        base = "/api/v1/agent/plan/" + prepared["run_id"]
        failed = self.clients["lin"].post(base + "/fail", json={"message": "桌面执行失败"})
        self.assertEqual(failed.status_code, 200, failed.text)
        self.assertEqual(failed.json(), {"ok": True})
        from server.store import database
        with database() as db:
            status = db.one("SELECT status FROM runs WHERE id=?", (prepared["run_id"],))["status"]
        self.assertEqual(status, "failed")
        completed = self.clients["lin"].post("/api/v1/agent/plan/complete", json={
            "run_id": prepared["run_id"], "text": "{}", "usage": {}, "model": "plan-test-model",
        })
        self.assertEqual(completed.status_code, 409, completed.text)

    def test_other_user_cannot_fail_agent_plan(self):
        prepared = self.prepare_agent_plan()
        response = self.clients["chen"].post(
            f"/api/v1/agent/plan/{prepared['run_id']}/fail", json={"message": "不属于当前用户"}
        )
        self.assertEqual(response.status_code, 404, response.text)
        completed = self.clients["lin"].post("/api/v1/agent/plan/complete", json={
            "run_id": prepared["run_id"], "text": json.dumps({"title": "归属校验", "subtitle": "", "sections": [
                {"heading": "结果", "points": ["仍可由所有者完成"]},
            ]}, ensure_ascii=False), "usage": {}, "model": "plan-test-model",
        })
        self.assertEqual(completed.status_code, 200, completed.text)

    def test_repeated_agent_plan_fail_returns_conflict(self):
        prepared = self.prepare_agent_plan()
        path = f"/api/v1/agent/plan/{prepared['run_id']}/fail"
        first = self.clients["lin"].post(path, json={"message": "第一次失败"})
        second = self.clients["lin"].post(path, json={"message": "重复失败"})
        self.assertEqual(first.status_code, 200, first.text)
        self.assertEqual(second.status_code, 409, second.text)

    def test_agent_plan_fail_rejects_wrong_assignment_paths(self):
        assignment = self.create_assignment()
        prepared = self.prepare_agent_plan(assignment)
        run_id = prepared["run_id"]
        unassigned_path = f"/api/v1/agent/plan/{run_id}/fail"
        wrong_assignment_path = f"/api/v1/assignments/wrong-key/sessions/{run_id}/fail"
        self.assertEqual(self.clients["lin"].post(unassigned_path, json={"message": "路径不符"}).status_code, 404)
        self.assertEqual(self.clients["lin"].post(wrong_assignment_path, json={"message": "路径不符"}).status_code, 404)
        matched = self.clients["lin"].post(
            f"/api/v1/assignments/{assignment}/sessions/{run_id}/fail", json={"message": "路径匹配"}
        )
        self.assertEqual(matched.status_code, 200, matched.text)

        unassigned = self.prepare_agent_plan()
        mismatch = self.clients["lin"].post(
            f"/api/v1/assignments/{assignment}/sessions/{unassigned['run_id']}/fail", json={"message": "没有任务"}
        )
        self.assertEqual(mismatch.status_code, 404, mismatch.text)

    def test_agent_plan_fail_hides_missing_or_non_desktop_runs(self):
        missing = self.clients["lin"].post(
            "/api/v1/agent/plan/not-a-run/fail", json={"message": "不存在"}
        )
        self.assertEqual(missing.status_code, 404, missing.text)
        prepared = self.prepare_agent_plan()
        from server.store import database
        with database(write=True) as db:
            db.execute("UPDATE runs SET engine='demo' WHERE id=?", (prepared["run_id"],))
        response = self.clients["lin"].post(
            f"/api/v1/agent/plan/{prepared['run_id']}/fail", json={"message": "engine 不匹配"}
        )
        self.assertEqual(response.status_code, 404, response.text)

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
