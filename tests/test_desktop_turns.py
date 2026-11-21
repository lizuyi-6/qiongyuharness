import json
import os
import tempfile
import unittest
import asyncio
import base64
import hashlib
import struct
import zlib
from uuid import uuid4
from io import BytesIO
from unittest.mock import patch

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
        cls.app = app
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

    def complete(self, key, prepared, actor="lin", text=None, usage=None, pipeline=None):
        value = text or json.dumps({
            "action": "export", "reply": "已整理演示稿", "kind": "pptx",
            "plan": {"title": "桌面执行验证", "subtitle": "", "sections": [
                {"heading": "目标", "points": ["核对接入方案"]},
                {"heading": "安排", "points": ["补充验证记录"]},
            ]},
        }, ensure_ascii=False)
        payload = {"text": value, "usage": usage or {"prompt_tokens": 13, "completion_tokens": 9, "total_tokens": 22},
                   "model": "desktop-test-model"}
        if pipeline is not None:
            payload["pipeline"] = pipeline
        return self.clients[actor].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/complete",
            json=payload,
        )

    def test_pipeline_config_is_authenticated_and_has_expected_shape(self):
        anonymous = TestClient(self.app).get("/api/v1/me/pipeline")
        self.assertEqual(anonymous.status_code, 401, anonymous.text)

        response = self.clients["lin"].get("/api/v1/me/pipeline")
        self.assertEqual(response.status_code, 200, response.text)
        value = response.json()
        self.assertIs(value["enabled"], True)
        self.assertEqual(set(value["roles"]), {"understand", "execute", "accept", "review", "light"})
        self.assertEqual(value["limits"], {"max_exec_rounds": 2, "deadline_ms": 270000})
        self.assertEqual(set(value["prompts"]), {"understand", "accept", "review", "revise"})
        self.assertEqual(len(value["version"]), 12)

    def test_prepare_injects_matched_writing_preset(self):
        key = self.conversation()
        response = self.clients["lin"].post(f"/api/v1/conversations/{key}/turns/prepare", json={
            "id": str(uuid4()), "message": "帮我把今天的会议笔记整理成纪要", "mode": "docx"
        })
        self.assertEqual(response.status_code, 200, response.text)
        prepared = response.json()
        self.assertIn("套用写作规范：会议纪要", prepared["steps"])
        self.assertIn("## 写作规范：会议纪要", prepared["system"])
        self.assertIn("## Word 导出规则", prepared["system"])
        payload = json.loads(prepared["prompt"])
        self.assertEqual(payload["写作规范"]["类型"], "会议纪要")
        self.assertEqual(payload["写作规范"]["最低分级"], "standard")
        self.assertIn("会议决定与讨论要点分开写", payload["写作规范"]["验收要点"])

        plain_key = self.conversation()
        plain = self.clients["lin"].post(f"/api/v1/conversations/{plain_key}/turns/prepare", json={
            "id": str(uuid4()), "message": "你好", "mode": "reply"
        }).json()
        self.assertFalse(any(step.startswith("套用写作规范") for step in plain["steps"]))
        self.assertNotIn("## 写作规范：", plain["system"])
        self.assertNotIn("写作规范", json.loads(plain["prompt"]))

    def test_plan_prepare_injects_preset_for_high_stakes_document(self):
        response = self.clients["lin"].post("/api/v1/agent/plan/prepare", json={
            "kind": "docx", "title": "技术服务合同", "instructions": "起草一份技术服务合同，我方是甲方", "pages": 8
        })
        self.assertEqual(response.status_code, 200, response.text)
        prepared = response.json()
        self.assertIn("套用写作规范：协议起草", prepared["steps"])
        self.assertIn("## 写作规范：协议起草", prepared["system"])
        self.assertIn('{"title":"标题"', prepared["system"])
        agenda = json.loads(prepared["prompt"].split("\n", 1)[1])
        self.assertEqual(agenda["写作规范"]["最低分级"], "critical")

    def test_invalid_pipeline_environment_value_fails_app_startup(self):
        async def enter_lifespan():
            async with self.app.router.lifespan_context(self.app):
                return None

        for values, expected in [
            ({"QO_PIPELINE_ENABLED": "yes"}, "QO_PIPELINE_ENABLED"),
            ({"QO_PIPELINE_ENABLED": "1", "QO_PIPELINE_MODEL_ACCEPT": "bad model"}, "QO_PIPELINE_MODEL_ACCEPT"),
        ]:
            with self.subTest(values=values), patch.dict(os.environ, values):
                with self.assertRaisesRegex(RuntimeError, expected):
                    asyncio.run(enter_lifespan())

    def test_pipeline_trace_model_rejects_invalid_values_and_preserves_legacy_shape(self):
        from pydantic import ValidationError
        from server.conversations import TurnCompletion

        stage = {"role": "execute", "model": "step-5-preview", "outcome": "ok",
                 "usage": {"prompt_tokens": 4, "private": "drop"}}
        base = {"tier": "standard", "verdict": "pass", "review": "not_needed", "issues": [],
                "stages": [stage], "rounds": 1, "version": "0123456789ab"}
        completion = TurnCompletion(text='{"action":"reply","reply":"ok"}', usage={}, model="desktop-test",
                                    pipeline=base)
        self.assertEqual(completion.pipeline.stages[0].usage, {"prompt_tokens": 4})
        self.assertIsNone(TurnCompletion(text="{}", usage={}, model="legacy-client").pipeline)

        with self.assertRaises(ValidationError):
            TurnCompletion(text="{}", usage={}, model="x", pipeline={**base, "verdict": "skipped"})
        with self.assertRaises(ValidationError):
            TurnCompletion(text="{}", usage={}, model="x", pipeline={**base, "stages": [{**stage, "role": "unknown"}]})
        with self.assertRaises(ValidationError):
            TurnCompletion(text="{}", usage={}, model="x", pipeline={**base, "issues": ["x" * 301]})

    def test_pipeline_trace_is_stored_in_turn_result(self):
        key = self.conversation()
        prepared = self.prepare(key)
        trace = {
            "tier": "standard", "verdict": "pass", "review": "not_needed", "issues": [],
            "stages": [{"role": "understand", "model": "glm-5.3", "outcome": "ok",
                        "usage": {"prompt_tokens": 5, "completion_tokens": 2, "total_tokens": 7}}],
            "rounds": 1, "version": "0123456789ab",
        }
        response = self.complete(key, prepared, pipeline=trace)
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json()["pipeline"], trace)
        from server.store import database, unpack
        with database() as db:
            stored = unpack(db.one("SELECT result FROM office_turns WHERE id=?", (prepared["run_id"],))["result"])
        self.assertEqual(stored["pipeline"], trace)

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
        # 0.2.14 起，不以 { 开头的纯文本按兜底 reply 交付，不再 502。
        accepted = self.complete(key, prepared, text="not a JSON decision")
        self.assertEqual(accepted.status_code, 200, accepted.text)
        self.assertEqual(accepted.json()["status"], "completed")
        self.assertEqual(accepted.json()["reply"], "not a JSON decision")
        again = self.complete(key, prepared)
        self.assertEqual(again.status_code, 409, again.text)
        # 以 { 开头但无法解析的残缺 JSON 仍走 502 失败路径，且本轮不能再补交。
        other = self.conversation()
        other_prepared = self.prepare(other)
        rejected = self.complete(other, other_prepared, text='{"action": "broken')
        self.assertEqual(rejected.status_code, 502, rejected.text)
        self.assertEqual(self.complete(other, other_prepared).status_code, 409)
        turn = self.clients["lin"].get(f"/api/v1/conversations/{other}").json()["turns"][-1]
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

    def image_png(self, width=1024, height=1024, trailing=b""):
        header = b"\x89PNG\r\n\x1a\n" + struct.pack(">I", 13) + b"IHDR" + struct.pack(">II", width, height)
        return header + trailing

    @staticmethod
    def real_png(width=1, height=1):
        # 带 CRC 与 IDAT/IEND 的最小合法灰度 PNG：能通过上传路径的解码校验（PIL 可用时）。
        def chunk(tag, data):
            return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

        ihdr = struct.pack(">IIBBBBB", width, height, 8, 0, 0, 0, 0)
        raw = b"".join(b"\x00" + b"\x00" * width for _ in range(height))
        return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")

    def image_payload(self, *, width=1024, height=1024, model="gpt-image-2.6-flare", size="1024x1024",
                      quality="medium", prompt="蓝色山峦插图", png=None):
        return {"png_base64": base64.b64encode(png if png is not None else self.image_png(width, height)).decode(),
                "width": width, "height": height, "model": model, "size": size, "quality": quality,
                "prompt": prompt, "usage": {"input_tokens": 7, "output_tokens": 9,
                    "input_tokens_details": {"image_tokens": 3, "private": 5}, "private": "drop"}}

    def prepare_image(self, key, actor="lin", **overrides):
        body = {"prompt": "蓝色山峦插图", "model": "gpt-image-2.6-flare", "size": "1024x1024", "quality": "medium"}
        body.update(overrides)
        return self.clients[actor].post(f"/api/v1/conversations/{key}/images/prepare", json=body)

    def test_image_button_prepare_strictly_validates_and_checks_owner_and_mode(self):
        key = self.conversation()
        invalid = [
            {"prompt": " ", "model": "gpt-image-2.6-flare", "size": "1024x1024", "quality": "medium"},
            {"prompt": "x", "model": "chat-model", "size": "1024x1024", "quality": "medium"},
            {"prompt": "x", "model": "gpt-image-2.6-flare", "size": "512x512", "quality": "medium"},
            {"prompt": "x", "model": "gpt-image-2.6-flare", "size": "1024x1024", "quality": "ultra"},
            {"prompt": "x", "model": "gpt-image-2.6-flare", "size": "1024x1024", "quality": "medium", "extra": True},
        ]
        invalid_fields = ["prompt", "model", "size", "quality", "extra"]
        for body, field in zip(invalid, invalid_fields):
            response = self.clients["lin"].post(f"/api/v1/conversations/{key}/images/prepare", json=body)
            self.assertEqual(response.status_code, 422, response.text)
            self.assertTrue(any(item.get("loc", [])[-1:] == [field] for item in response.json()["detail"]), response.text)
        other = self.prepare_image(key, actor="chen")
        self.assertEqual(other.status_code, 404, other.text)
        self.assertEqual(other.json()["detail"], "会话不存在")
        previous = self.gateway.MODE
        try:
            self.gateway.MODE = "demo"
            disabled = self.prepare_image(key)
            self.assertEqual(disabled.status_code, 409, disabled.text)
            self.assertEqual(disabled.json()["detail"], "桌面执行模式未启用")
        finally:
            self.gateway.MODE = previous

    def test_button_image_complete_checks_png_params_and_persists_digest_and_rejects_repeat(self):
        from server.store import database, unpack
        key = self.conversation()
        prepared = self.prepare_image(key).json()
        path = f"/api/v1/conversations/{key}/images/{prepared['run_id']}/complete"
        wrong_model_payload = self.image_payload(model="gpt-image-2.6-sunburst")
        wrong_model_payload.pop("prompt")
        wrong_model = self.clients["lin"].post(path, json=wrong_model_payload)
        self.assertEqual(wrong_model.status_code, 422, wrong_model.text)
        self.assertEqual(wrong_model.json()["detail"], "生图参数与准备请求不一致")
        mismatch_payload = self.image_payload(width=1536, height=1024,
            size="1536x1024", png=self.image_png(1024, 1024))
        mismatch_payload.pop("prompt")
        mismatch = self.clients["lin"].post(path, json=mismatch_payload)
        self.assertEqual(mismatch.status_code, 422, mismatch.text)
        self.assertEqual(mismatch.json()["detail"], "生图返回内容无效")
        non_png_payload = self.image_payload(png=b"not a PNG")
        non_png_payload.pop("prompt")
        non_png = self.clients["lin"].post(path, json=non_png_payload)
        self.assertEqual(non_png.status_code, 422, non_png.text)
        self.assertEqual(non_png.json()["detail"], "生图返回内容无效")
        oversized_png = self.image_png() + b"x" * (10 * 1024 * 1024)
        oversized_payload = self.image_payload(png=oversized_png)
        oversized_payload.pop("prompt")
        too_large = self.clients["lin"].post(path, json=oversized_payload)
        self.assertEqual(too_large.status_code, 422, too_large.text)
        self.assertTrue(any(item.get("loc", [])[-1:] == ["png_base64"] for item in too_large.json()["detail"]),
                        too_large.text)
        valid = self.image_payload()
        result = self.clients["lin"].post(path, json={key: value for key, value in valid.items() if key != "prompt"})
        self.assertEqual(result.status_code, 200, result.text)
        turn = result.json()
        self.assertEqual(turn["status"], "completed")
        self.assertEqual(len(turn["images"]), 1)
        stored_image = turn["images"][0]
        self.assertEqual(stored_image["source"], "user")
        self.assertEqual(stored_image["prompt"], "蓝色山峦插图")
        self.assertEqual(stored_image["usage"], {"input_tokens": 7, "output_tokens": 9,
            "input_tokens_details": {"image_tokens": 3}})
        file_id = stored_image["file"]["id"]
        data = base64.b64decode(valid["png_base64"])
        with database() as db:
            artifact = db.one("SELECT extension,label,run_id,digest,data FROM artifacts WHERE id=?", (file_id,))
            image_audit = db.one("SELECT type,payload FROM events WHERE type='image.generated' AND payload LIKE ?",
                                 (f"%{prepared['run_id']}%",))
            usage = unpack(db.one("SELECT usage FROM runs WHERE id=?", (prepared["run_id"],))["usage"])
        self.assertEqual(artifact["extension"], "png")
        self.assertEqual(artifact["label"], "image")
        self.assertEqual(artifact["run_id"], prepared["run_id"])
        self.assertEqual(artifact["digest"], hashlib.sha256(data).hexdigest())
        self.assertEqual(bytes(artifact["data"]), data)
        self.assertEqual(image_audit["type"], "image.generated")
        self.assertNotIn("png_base64", image_audit["payload"])
        self.assertEqual(usage["model"], "gpt-image-2.6-flare")
        duplicate = self.clients["lin"].post(path, json={key: value for key, value in valid.items() if key != "prompt"})
        self.assertEqual(duplicate.status_code, 409, duplicate.text)
        self.assertEqual(duplicate.json()["detail"], "本轮生图已结束")

    def test_button_image_wrong_owner_and_fail_keep_lifecycle(self):
        key = self.conversation()
        prepared = self.prepare_image(key).json()
        base = f"/api/v1/conversations/{key}/images/{prepared['run_id']}"
        body = self.image_payload()
        body.pop("prompt")
        other = self.clients["chen"].post(base + "/complete", json=body)
        self.assertEqual(other.status_code, 404, other.text)
        self.assertEqual(other.json()["detail"], "会话不存在")
        failed = self.clients["lin"].post(base + "/fail", json={"message": "生成中断"})
        self.assertEqual(failed.status_code, 200, failed.text)
        history = self.clients["lin"].get(f"/api/v1/conversations/{key}").json()["turns"][-1]
        self.assertEqual(history["status"], "failed")
        self.assertEqual(history["images"], [])
        self.assertEqual(history["error"], "生成中断")
        duplicate = self.clients["lin"].post(base + "/fail", json={"message": "再次失败"})
        self.assertEqual(duplicate.status_code, 409, duplicate.text)
        self.assertEqual(duplicate.json()["detail"], "本轮生图已结束")

    def test_ai_image_uploads_append_without_cap_and_complete_preserves_them(self):
        key = self.conversation()
        prepared = self.prepare(key)
        run_path = f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/images"
        image_body = self.image_payload()
        first = self.clients["lin"].post(run_path, json=image_body)
        self.assertEqual(first.status_code, 200, first.text)
        first_image = first.json()
        self.assertEqual(first_image["source"], "ai")
        self.assertEqual(first_image["prompt"], "蓝色山峦插图")
        self.assertEqual(first_image["usage"], {"input_tokens": 7, "output_tokens": 9,
            "input_tokens_details": {"image_tokens": 3}})
        # 自由画布工作流按需生图，服务端不再限制每轮张数，序号依次递增。
        uploads = [first]
        for prompt in ("另一张插图", "第三张配图"):
            upload = self.clients["lin"].post(run_path, json=self.image_payload(prompt=prompt))
            self.assertEqual(upload.status_code, 200, upload.text)
            uploads.append(upload)
        self.assertEqual([upload.json()["file"]["name"] for upload in uploads],
                         ["AI配图-1.png", "AI配图-2.png", "AI配图-3.png"])
        complete = self.complete(key, prepared)
        self.assertEqual(complete.status_code, 200, complete.text)
        final = complete.json()
        self.assertEqual(final["status"], "completed")
        self.assertEqual([image["file"]["id"] for image in final["images"]],
                         [upload.json()["file"]["id"] for upload in uploads])
        self.assertEqual(final["reply"], "已整理演示稿")
        stale = self.prepare(key)
        self.assertEqual(stale["kind"], "conversation")
        self.assertEqual(stale["run_id"] != prepared["run_id"], True)

    def test_ai_images_survive_failed_turn_and_nonrunning_and_nonowner_reject(self):
        key = self.conversation()
        prepared = self.prepare(key)
        run_path = f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/images"
        other = self.clients["chen"].post(run_path, json=self.image_payload())
        self.assertEqual(other.status_code, 404, other.text)
        self.assertEqual(other.json()["detail"], "会话不存在")
        saved = self.clients["lin"].post(run_path, json=self.image_payload())
        self.assertEqual(saved.status_code, 200, saved.text)
        failed = self.clients["lin"].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/fail", json={"message": "模型后续失败"})
        self.assertEqual(failed.status_code, 200, failed.text)
        history = self.clients["lin"].get(f"/api/v1/conversations/{key}").json()["turns"][-1]
        self.assertEqual(history["status"], "failed")
        self.assertEqual(len(history["images"]), 1)
        self.assertEqual(history["images"][0]["file"]["id"], saved.json()["file"]["id"])
        after = self.clients["lin"].post(run_path, json=self.image_payload())
        self.assertEqual(after.status_code, 409, after.text)
        self.assertEqual(after.json()["detail"], "本轮执行已结束")

    def test_image_view_is_owner_only_and_uploads_validate_png_content(self):
        key = self.conversation()
        prepared = self.prepare_image(key).json()
        payload = self.image_payload()
        png_bytes = base64.b64decode(payload["png_base64"])
        payload.pop("prompt")
        completed = self.clients["lin"].post(
            f"/api/v1/conversations/{key}/images/{prepared['run_id']}/complete",
            json={name: value for name, value in payload.items() if name != "prompt"})
        self.assertEqual(completed.status_code, 200, completed.text)
        image_id = completed.json()["images"][0]["file"]["id"]
        shown = self.clients["lin"].get(f"/api/v1/files/{image_id}/image")
        self.assertEqual(shown.status_code, 200, shown.text)
        self.assertEqual(shown.content, png_bytes)
        self.assertEqual(shown.headers["content-type"], "image/png")
        self.assertEqual(shown.headers["x-content-type-options"], "nosniff")
        self.assertEqual(shown.headers["content-disposition"], "inline")
        self.assertEqual(shown.headers["cache-control"], "private, max-age=0")
        other = self.clients["chen"].get(f"/api/v1/files/{image_id}/image")
        self.assertEqual(other.status_code, 404, other.text)
        self.assertEqual(other.json()["detail"], "文件不存在或当前账号无权访问")
        # 0.2.21 起真 PNG 可作为普通文件上传，并经 /files/{id}/image 预览。
        real_png = self.real_png()
        uploaded = self.clients["lin"].post("/api/v1/files", files={"file": ("上传图.png", real_png, "image/png")})
        self.assertEqual(uploaded.status_code, 200, uploaded.text)
        uploaded_file = uploaded.json()
        self.assertEqual(uploaded_file["extension"], "png")
        self.assertEqual(uploaded_file["label"], "uploaded")
        served = self.clients["lin"].get(f"/api/v1/files/{uploaded_file['id']}/image")
        self.assertEqual(served.status_code, 200, served.text)
        self.assertEqual(served.content, real_png)
        # 无 PNG 文件头的伪造内容仍被拒绝。
        rejected = self.clients["lin"].post("/api/v1/files", files={"file": ("坏图.png", b"not a png", "image/png")})
        self.assertEqual(rejected.status_code, 422, rejected.text)
        self.assertEqual(rejected.json()["detail"], "PNG 图片格式无效")
        doc_turn = self.prepare(key)
        document = self.complete(key, doc_turn)
        self.assertEqual(document.status_code, 200, document.text)
        non_image = self.clients["lin"].get(f"/api/v1/files/{document.json()['file']['id']}/image")
        self.assertEqual(non_image.status_code, 404, non_image.text)
        self.assertEqual(non_image.json()["detail"], "图片不存在")

    def test_ai_image_upload_requires_desktop_mode_and_valid_payload(self):
        key = self.conversation()
        prepared = self.prepare(key)
        path = f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/images"
        invalid = self.image_payload(png=b"not PNG")
        rejected = self.clients["lin"].post(path, json=invalid)
        self.assertEqual(rejected.status_code, 422, rejected.text)
        self.assertEqual(rejected.json()["detail"], "生图返回内容无效")
        previous = self.gateway.MODE
        try:
            self.gateway.MODE = "demo"
            disabled = self.clients["lin"].post(path, json=self.image_payload())
            self.assertEqual(disabled.status_code, 409, disabled.text)
            self.assertEqual(disabled.json()["detail"], "桌面执行模式未启用")
        finally:
            self.gateway.MODE = previous
        failed = self.clients["lin"].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/fail", json={"message": "cleanup"})
        self.assertEqual(failed.status_code, 200, failed.text)

    def test_oversized_recent_deck_is_analyzed_before_injection(self):
        # 线上工单根因：整套自由画布 deck（数百 KB）原样注入「最近文稿」撑爆模型上下文。
        key = self.conversation()
        first = self.prepare(key)
        slides = []
        for page in range(1, 31):
            slides.append({"title": f"第{page}页 议题", "notes": "演讲备注" * 40,
                           "elements": [{"type": "text", "x": 0.7, "y": 0.5 + index * 0.32, "w": 11.9, "h": 0.28,
                                         "size": 14, "color": "#333333",
                                         "text": f"第{page}页要点{index} " + "内容" * 50}
                                        for index in range(20)]})
        deck = {"title": "巨幅演示", "subtitle": "上下文压力测试",
                "design": {"palette": {"bg": "#ffffff"}, "font": "Microsoft YaHei"}, "slides": slides}
        completed = self.complete(key, first, text=json.dumps(
            {"action": "export", "reply": "已生成巨幅演示", "kind": "pptx", "deck": deck}, ensure_ascii=False))
        self.assertEqual(completed.status_code, 200, completed.text)
        second = self.prepare(key, message="修复这一版：加上图表")
        self.assertLess(len(second["prompt"]), 90_000, "巨幅 deck 注入前必须先压缩为结构摘要")
        self.assertIn("压缩说明", second["prompt"])
        self.assertIn("第30页 议题", second["prompt"])
        self.assertIn("第1页要点0", second["prompt"])
        self.clients["lin"].post(f"/api/v1/conversations/{key}/turns/{second['run_id']}/fail",
                                 json={"message": "继续下一轮"})
        # 小文稿不压缩：完整 plan 原样注入。
        third_prepared = self.prepare(key, message="换成两页简版")
        third = self.complete(key, third_prepared, text=json.dumps(
            {"action": "export", "reply": "已精简", "kind": "pptx",
             "deck": {"title": "简版", "subtitle": "", "slides": [
                 {"title": "唯一页", "notes": "", "elements": [
                     {"type": "text", "x": 0.7, "y": 0.5, "w": 11.9, "h": 0.6, "size": 24,
                      "color": "#111111", "text": "简版唯一要点"}]}]}}, ensure_ascii=False))
        self.assertEqual(third.status_code, 200, third.text)
        fourth = self.prepare(key, message="再改改标题")
        self.assertNotIn("压缩说明", fourth["prompt"])
        self.assertIn("简版", fourth["prompt"])

    def test_overlong_docx_point_is_truncated_and_saved(self):
        # 线上工单：模型单条要点超过 600 字时整轮 422，已生成已计费的结果被毁。
        key = self.conversation()
        response = self.clients["lin"].post(f"/api/v1/conversations/{key}/turns/prepare", json={
            "id": str(uuid4()), "message": "把商业计划书写成 Word 文档", "mode": "docx"})
        self.assertEqual(response.status_code, 200, response.text)
        prepared = response.json()
        long_point = "超长要点：" + "市场规模持续增长的原因描述 " * 50
        self.assertGreater(len(long_point), 600)
        text = json.dumps({"action": "export", "reply": "已生成商业计划书", "kind": "docx",
                           "plan": {"title": "商业计划书", "subtitle": "", "sections": [
                               {"heading": "市场分析", "points": ["标准要点", long_point, "   "]}]}},
                          ensure_ascii=False)
        completed = self.clients["lin"].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/complete",
            json={"text": text, "usage": {"total_tokens": 9}, "model": "desktop-test-model"})
        self.assertEqual(completed.status_code, 200, completed.text)
        body = completed.json()
        self.assertEqual(body["kind"], "docx")
        from server.store import database, unpack
        with database() as db:
            stored = unpack(db.one("SELECT result FROM office_turns WHERE id=?",
                                    (prepared["run_id"],))["result"])
        points = stored["plan"]["sections"][0]["points"]
        self.assertEqual(len(points), 2, "空白要点应被剔除而不是整轮拒绝")
        self.assertLessEqual(len(points[1]), 600)
        self.assertTrue(points[1].startswith("超长要点"), "截断应保留要点开头")

    def test_oversized_freeform_elements_render_and_save(self):
        # 线上工单续：自由画布超长文本框（>2000）与表格格（>60/>120）在渲染时截断，而不是 422 毁掉整轮。
        key = self.conversation()
        prepared = self.prepare(key, message="把商业计划书做成 PPT")
        deck = {"title": "商业计划书 PPT", "subtitle": "", "slides": [
            {"title": "对比页", "notes": "超长备注：" + "本页讲述的核心论据与数据来源说明 " * 120,
             "elements": [
                {"type": "text", "x": 0.7, "y": 0.5, "w": 11.9, "h": 1.2, "size": 14,
                 "text": "超长文本框：" + "详细市场规模分析叙述 " * 180},
                {"type": "table", "x": 0.7, "y": 2.0, "w": 11.9, "h": 3.0,
                 "columns": ["超长表头" + "列名" * 30, "B", "C"],
                 "rows": [["超长单元格" + "内容描述" * 40, "x", "y"], ["a", "b", "c"]]},
            ]}]}
        text = json.dumps({"action": "export", "reply": "已生成对比页", "kind": "pptx", "deck": deck},
                          ensure_ascii=False)
        completed = self.complete(key, prepared, text=text)
        self.assertEqual(completed.status_code, 200, completed.text)
        body = completed.json()
        self.assertEqual(body["kind"], "pptx")
        download = self.clients["lin"].get(f"/api/v1/files/{body['file']['id']}/download")
        self.assertEqual(download.status_code, 200, download.text)
        self.assertEqual(len(Presentation(BytesIO(download.content)).slides), 1)

    def test_overlong_docx_table_cells_are_truncated_and_saved(self):
        # 线上工单续：docx 表格超长表头（>40）/超长格（>80）/多写格子一律截断，不再 422 毁轮。
        key = self.conversation()
        response = self.clients["lin"].post(f"/api/v1/conversations/{key}/turns/prepare", json={
            "id": str(uuid4()), "message": "把商业计划书写成 Word 文档", "mode": "docx"})
        self.assertEqual(response.status_code, 200, response.text)
        prepared = response.json()
        text = json.dumps({"action": "export", "reply": "已生成", "kind": "docx",
                           "plan": {"title": "商业计划书", "subtitle": "", "sections": [
                               {"heading": "对比", "points": ["要点"], "table": {
                                   "columns": ["超长表头" + "列名" * 30, "B"],
                                   "rows": [["超长单元格" + "内容" * 60, "v", "多余格"], ["a", "b"]]}}]}},
                          ensure_ascii=False)
        completed = self.clients["lin"].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/complete",
            json={"text": text, "usage": {"total_tokens": 9}, "model": "desktop-test-model"})
        self.assertEqual(completed.status_code, 200, completed.text)
        from server.store import database, unpack
        with database() as db:
            stored = unpack(db.one("SELECT result FROM office_turns WHERE id=?",
                                    (prepared["run_id"],))["result"])
        table = stored["plan"]["sections"][0]["table"]
        self.assertLessEqual(len(table["columns"][0]), 40)
        self.assertTrue(table["columns"][0].startswith("超长表头"))
        self.assertEqual(len(table["rows"][0]), 2, "多写的格子应被丢弃而不是拒绝整轮")
        self.assertLessEqual(len(table["rows"][0][0]), 80)
        self.assertEqual(table["rows"][1], ["a", "b"])

    def test_turn_image_upload_is_idempotent_by_image_id(self):
        # 线上工单：上行被掐后客户端重试同一张已计费配图，服务端按 image_id 去重，不得重复入库。
        key = self.conversation()
        prepared = self.prepare(key)
        run_path = f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/images"
        payload = self.image_payload()
        payload["image_id"] = "retry-img-001"
        first = self.clients["lin"].post(run_path, json=payload)
        self.assertEqual(first.status_code, 200, first.text)
        second = self.clients["lin"].post(run_path, json=payload)
        self.assertEqual(second.status_code, 200, second.text)
        self.assertEqual(second.json()["file"]["id"], first.json()["file"]["id"], "重试应返回首次入库结果")
        self.assertEqual(second.json()["image_id"], "retry-img-001")
        history = self.clients["lin"].get(f"/api/v1/conversations/{key}").json()["turns"][-1]
        self.assertEqual(len(history["images"]), 1, "同一 image_id 不得重复计张")

    def test_overlong_deck_title_and_session_title_do_not_422_rollback(self):
        # 2026-11-21 线上 Turn 79b9518e 复现：PPT 标题 35 字（合法上限 110 字）在保存末尾
        # 自动命名会话时命中 28 字上限导致 422 事务回滚，修复后应截断到 <=28 字并正常保存。
        key = self.conversation()
        prepared = self.prepare(key, message="帮我生成挑战杯路演 PPT")
        long_title = "AetherGuard 边缘智能零信任工业网络安全网关挑战杯路演演示文稿全案"
        self.assertGreater(len(long_title), 28)
        deck = {
            "title": long_title,
            "subtitle": "副标题",
            "slides": [{"title": "封面", "elements": [{"type": "text", "x": 1, "y": 1, "w": 10, "h": 1, "text": "封面内容"}]}],
        }
        text = json.dumps({
            "action": "export",
            "reply": "已生成",
            "session_title": long_title,
            "kind": "pptx",
            "deck": deck,
        }, ensure_ascii=False)
        completed = self.complete(key, prepared, text=text)
        self.assertEqual(completed.status_code, 200, completed.text)
        conv = self.clients["lin"].get(f"/api/v1/conversations/{key}").json()
        self.assertLessEqual(len(conv["title"]), 28)
        self.assertTrue(conv["title"].startswith("AetherGuard"))

    def test_fail_turn_with_overlong_step_and_feedback_server_version(self):
        key = self.conversation()
        prepared = self.prepare(key)
        long_step = "DeepSeek 搜索：" + "超长检索关键词 " * 60
        failed = self.clients["lin"].post(
            f"/api/v1/conversations/{key}/turns/{prepared['run_id']}/fail",
            json={"message": "失败原因：" + "x" * 100, "steps": [long_step]},
        )
        self.assertEqual(failed.status_code, 200, failed.text)

        cfg = self.clients["lin"].get("/api/v1/config").json()
        report = self.clients["lin"].post("/api/v1/feedback", json={
            "category": "ai",
            "title": "版本一致性校验",
            "description": "检查 server_version 与 /api/v1/config 一致",
        })
        self.assertEqual(report.status_code, 200, report.text)
        self.assertEqual(report.json()["server_version"], cfg["version"])


if __name__ == "__main__":
    unittest.main()
