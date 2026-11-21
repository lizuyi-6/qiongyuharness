import json
import os
import sqlite3
import tempfile
import unittest
from datetime import UTC, datetime, timedelta
from pathlib import Path
from unittest.mock import patch
from urllib.parse import parse_qs, urlsplit

import httpx
import jwt
from cryptography.hazmat.primitives.asymmetric import rsa
from fastapi.testclient import TestClient

IMPORT_DATA = tempfile.TemporaryDirectory()
os.environ.update(QO_DEMO="1", QO_ENGINE="demo", QO_DATA_DIR=IMPORT_DATA.name)

from server import identity, store, sub2api
from server.app import app


class ZaochangAuthTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.previous_data = store.DATA
        cls.previous_demo = store.DEMO
        cls.previous_roster = os.environ.get("QO_USERS_FILE")
        cls.previous_open_enrollment = os.environ.get("QO_OPEN_ENROLLMENT")
        cls.data = Path(cls.temp.name) / "data"
        cls.data.mkdir()
        cls.roster_path = Path(cls.temp.name) / "users.json"
        cls.roster_path.write_text(json.dumps([
            {"email": "Admin@Example.com", "name": "管理员", "department": "运营", "role": "admin"},
            {"email": "worker@example.com", "name": "经办员工", "department": "产品", "role": "employee"},
            {"email": "unverified@example.com", "name": "未验证用户", "department": "产品", "role": "employee"},
            {"email": "reviewer@example.com", "name": "审核员", "department": "运营", "role": "reviewer"},
            {"email": "blocked@example.com", "name": "已停用用户", "department": "运营", "role": "employee", "enabled": False},
            {"email": "manager@example.com", "name": "部门主管", "department": "研发", "role": "manager"},
        ]), encoding="utf-8")
        store.DATA = cls.data
        store.DEMO = False
        os.environ["QO_USERS_FILE"] = str(cls.roster_path)
        os.environ["QO_OPEN_ENROLLMENT"] = "0"
        store.initialize()
        cls.issuer = "https://identity.test"
        cls.client_id = "office-test"
        cls.signing_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        cls.jwk = json.loads(jwt.algorithms.RSAAlgorithm.to_jwk(cls.signing_key.public_key()))
        cls.jwk.update({"kid": "zaochang-test", "use": "sig", "alg": "RS256"})

    @classmethod
    def tearDownClass(cls):
        store.DATA = cls.previous_data
        store.DEMO = cls.previous_demo
        if cls.previous_roster is None:
            os.environ.pop("QO_USERS_FILE", None)
        else:
            os.environ["QO_USERS_FILE"] = cls.previous_roster
        if cls.previous_open_enrollment is None:
            os.environ.pop("QO_OPEN_ENROLLMENT", None)
        else:
            os.environ["QO_OPEN_ENROLLMENT"] = cls.previous_open_enrollment
        cls.temp.cleanup()

    def setUp(self):
        store.initialize()

    def user_id(self, email):
        import hashlib
        return "u_" + hashlib.sha256(email.lower().encode("utf-8")).hexdigest()[:24]

    def oidc_login(self, subject, email=None, email_verified=None, *, code="login", desktop=False, name=None):
        action = {}
        real_client = httpx.AsyncClient

        def endpoint(request):
            if request.url.path == "/.well-known/openid-configuration":
                return httpx.Response(200, json={
                    "issuer": self.issuer,
                    "authorization_endpoint": self.issuer + "/authorize",
                    "token_endpoint": self.issuer + "/token",
                    "jwks_uri": self.issuer + "/keys",
                })
            if request.url.path == "/keys":
                return httpx.Response(200, json={"keys": [self.jwk]})
            if request.url.path == "/token":
                form = parse_qs(request.content.decode())
                claims = {
                    "iss": self.issuer,
                    "sub": subject,
                    "aud": self.client_id,
                    "iat": datetime.now(UTC),
                    "exp": datetime.now(UTC) + timedelta(minutes=4),
                    "nonce": action["nonce"][0],
                }
                if email is not None:
                    claims["email"] = email
                if email_verified is not None:
                    claims["email_verified"] = email_verified
                if name is not None:
                    claims["name"] = name
                token = jwt.encode(claims, self.signing_key, algorithm="RS256", headers={"kid": self.jwk["kid"]})
                return httpx.Response(200, json={"id_token": token})
            return httpx.Response(404)

        params = {"desktop": identity.challenge("v" * 64), "port": "43291"} if desktop else {}
        with (
            patch.object(identity, "ISSUER", self.issuer),
            patch.object(identity, "CLIENT", self.client_id),
            patch.object(identity, "CLIENT_SECRET", ""),
            patch.object(identity, "ORIGIN", "https://office.test"),
            patch.object(identity, "SECURE", False),
            patch.object(identity.httpx, "AsyncClient", side_effect=lambda **kw: real_client(transport=httpx.MockTransport(endpoint), **kw)),
        ):
            with TestClient(app, follow_redirects=False) as client:
                start = client.get("/auth/start", params=params)
                self.assertEqual(start.status_code, 303, start.text)
                action.update(parse_qs(urlsplit(start.headers["location"]).query))
                response = client.get("/auth/callback", params={"state": action["state"][0], "code": code})
                return response, client, action

    def test_email_bind_then_subject_login_and_denials(self):
        with store.database() as db:
            admin = db.one("SELECT id,email FROM users WHERE email=?", ("admin@example.com",))
            self.assertEqual(admin["id"], self.user_id("admin@example.com"))
            self.assertEqual(admin["email"], "admin@example.com")
            listed = db.one("SELECT name,department,role,enabled,source FROM users WHERE email=?", ("manager@example.com",))
            self.assertEqual((listed["name"], listed["department"], listed["role"], listed["enabled"], listed["source"]),
                             ("部门主管", "研发", "manager", 1, "roster"))
        first, _, action = self.oidc_login("subject-worker", "WORKER@example.com", True, code="first")
        self.assertEqual(first.status_code, 303, first.text)
        self.assertEqual(action["scope"], ["openid profile email"])
        worker_id = self.user_id("worker@example.com")
        with store.database() as db:
            self.assertEqual(db.one("SELECT subject FROM users WHERE id=?", (worker_id,))["subject"], "subject-worker")
            self.assertEqual(db.one("SELECT COUNT(*) AS count FROM events WHERE actor_id=? AND type='auth.subject_bound'", (worker_id,))["count"], 1)

        second, _, _ = self.oidc_login("subject-worker", None, False, code="by-sub")
        self.assertEqual(second.status_code, 303, second.text)
        conflict, _, _ = self.oidc_login("changed-subject", "worker@example.com", True, code="changed-sub")
        self.assertEqual(conflict.status_code, 403)
        self.assertEqual(conflict.json()["detail"], "此邮箱已绑定其他造场账号")
        unverified, _, _ = self.oidc_login("subject-unverified", "unverified@example.com", False, code="unverified")
        self.assertEqual(unverified.status_code, 403)
        outside, _, _ = self.oidc_login("subject-outside", "outside@example.com", True, code="outside")
        self.assertEqual(outside.status_code, 403)
        missing, _, _ = self.oidc_login("subject-missing", None, None, code="missing")
        self.assertEqual(missing.status_code, 403)
        with store.database() as db:
            self.assertIsNone(db.one("SELECT subject FROM users WHERE id=?", (self.user_id("unverified@example.com"),))["subject"])

    def test_open_enrollment_creates_employee_and_keeps_subject_binding(self):
        with patch.dict(os.environ, {"QO_OPEN_ENROLLMENT": "1"}, clear=False):
            response, _, _ = self.oidc_login(
                "auto-subject", " New.Hire@Example.com ", True, code="auto-enroll", name="新入职员工"
            )
            self.assertEqual(response.status_code, 303, response.text)
            user_id = self.user_id("new.hire@example.com")
            with store.database() as db:
                user = db.one("SELECT id,email,subject,name,department,role,enabled,source FROM users WHERE id=?", (user_id,))
                self.assertEqual(tuple(user), (
                    user_id, "new.hire@example.com", "auto-subject", "新入职员工", "运营", "employee", 1, "auto"
                ))
                event = db.one("SELECT type,payload FROM events WHERE actor_id=? AND type='user.auto_enrolled'", (user_id,))
                self.assertEqual(event["type"], "user.auto_enrolled")
                payload = store.unpack(event["payload"])
                self.assertEqual(payload["user_id"], user_id)
                self.assertEqual(payload["email_sha256"], identity.digest("new.hire@example.com"))
                self.assertNotIn("new.hire@example.com", event["payload"])
            store.initialize()
            with store.database() as db:
                self.assertEqual(db.one("SELECT enabled FROM users WHERE id=?", (user_id,))["enabled"], 1)

            fallback, _, _ = self.oidc_login("fallback-subject", "fallback@example.com", True, code="auto-name")
            self.assertEqual(fallback.status_code, 303, fallback.text)
            with store.database() as db:
                fallback_user = db.one("SELECT name,role,department FROM users WHERE email=?", ("fallback@example.com",))
                self.assertEqual(tuple(fallback_user), ("fallback", "employee", "运营"))

            conflict, _, _ = self.oidc_login("changed-auto-subject", "new.hire@example.com", True, code="auto-conflict")
            self.assertEqual(conflict.status_code, 403)
            self.assertEqual(conflict.json()["detail"], "此邮箱已绑定其他造场账号")
            with store.database() as db:
                self.assertEqual(db.one("SELECT subject FROM users WHERE id=?", (user_id,))["subject"], "auto-subject")

    def test_open_enrollment_requires_strictly_verified_email(self):
        with patch.dict(os.environ, {"QO_OPEN_ENROLLMENT": "1"}, clear=False):
            for index, verified in enumerate((False, None, "true", 1)):
                email = f"unverified-{index}@example.com"
                response, _, _ = self.oidc_login(
                    f"unverified-subject-{index}", email, verified, code=f"unverified-auto-{index}"
                )
                self.assertEqual(response.status_code, 403)
                self.assertEqual(response.json()["detail"], "造场账号未提供已验证邮箱，无法自动注册工作台账号")
                with store.database() as db:
                    self.assertIsNone(db.one("SELECT id FROM users WHERE email=?", (email,)))
                    self.assertEqual(db.one("SELECT COUNT(*) AS count FROM events WHERE type='user.auto_enrolled' AND payload LIKE ?",
                                            (f"%{identity.digest(email)}%",))["count"], 0)

    def test_roster_enabled_false_blocks_open_enrollment_and_clears_credentials(self):
        blocked_id = self.user_id("blocked@example.com")
        with store.database(write=True) as db:
            session = identity.make_session(db, blocked_id)
            db.execute(
                "INSERT INTO desktop_auth_codes(code_hash,challenge,user_id,expires) VALUES(?,?,?,?)",
                (identity.digest("unused-code"), "challenge", blocked_id, store.instant(datetime.now(UTC) + timedelta(minutes=1))),
            )
        with patch.dict(os.environ, {"QO_OPEN_ENROLLMENT": "1"}, clear=False):
            store.initialize()
            with store.database() as db:
                user = db.one("SELECT enabled,source,roster_blocked FROM users WHERE id=?", (blocked_id,))
                self.assertEqual(tuple(user), (0, "roster", 1))
                self.assertIsNone(db.one("SELECT token FROM sessions WHERE token=?", (identity.digest(session),)))
                self.assertEqual(db.one("SELECT COUNT(*) AS count FROM desktop_auth_codes WHERE user_id=?", (blocked_id,))["count"], 0)
            response, _, _ = self.oidc_login("blocked-subject", "blocked@example.com", True, code="blocked-member")
        self.assertEqual(response.status_code, 403)
        self.assertEqual(response.json()["detail"], "此邮箱已被管理员停用")
        with store.database() as db:
            self.assertEqual(db.one("SELECT COUNT(*) AS count FROM users WHERE email=?", ("blocked@example.com",))["count"], 1)

    def test_closed_enrollment_preserves_legacy_sync_disable_behavior(self):
        with patch.dict(os.environ, {"QO_OPEN_ENROLLMENT": "1"}, clear=False):
            enrolled, _, _ = self.oidc_login("closed-mode-subject", "closed-mode@example.com", True, code="closed-create")
            self.assertEqual(enrolled.status_code, 303, enrolled.text)
        user_id = self.user_id("closed-mode@example.com")
        with store.database(write=True) as db:
            session = identity.make_session(db, user_id)
        with patch.dict(os.environ, {"QO_OPEN_ENROLLMENT": "0"}, clear=False):
            store.initialize()
            with store.database() as db:
                disabled = db.one("SELECT enabled,source FROM users WHERE id=?", (user_id,))
                self.assertEqual(tuple(disabled), (0, "auto"))
                self.assertIsNone(db.one("SELECT token FROM sessions WHERE token=?", (identity.digest(session),)))
            denied, _, _ = self.oidc_login("closed-mode-subject", "closed-mode@example.com", True, code="closed-deny")
        self.assertEqual(denied.status_code, 403)
        self.assertEqual(denied.json()["detail"], "此账号尚未开通工作台权限")

    def test_open_enrollment_setting_rejects_values_other_than_zero_or_one(self):
        with patch.dict(os.environ, {}, clear=True):
            self.assertIs(store.open_enrollment_enabled(), False)
        for value in ("true", "", " 1", "2"):
            with patch.dict(os.environ, {"QO_OPEN_ENROLLMENT": value}, clear=False):
                with self.assertRaisesRegex(RuntimeError, "QO_OPEN_ENROLLMENT must be exactly '0' or '1'"):
                    identity.validate_identity_config()

    def make_session(self, email):
        with store.database(write=True) as db:
            return identity.make_session(db, self.user_id(email))

    def call_model_credential(self, email, handler, extra_env=None):
        real_client = httpx.AsyncClient
        config_env = {
            "QO_SUB2API_URL": "https://aethernettech.dev",
            "QO_SUB2API_PROVISION_KEY": "adm_sc_test_restricted-provision-key",
            "QO_SUB2API_DEFAULT_GROUP_ID": "7",
            "QO_SUB2API_KEY_NAME": "qiongyu-office",
            "QO_SUB2API_RATE_LIMIT_5H_USD": "20",
            "QO_SUB2API_RATE_LIMIT_1D_USD": "",
            "QO_SUB2API_RATE_LIMIT_7D_USD": "",
        }
        config_env.update(extra_env or {})
        with (
            patch.object(sub2api, "DEMO", False),
            patch.dict(os.environ, config_env),
            patch.object(sub2api.httpx, "AsyncClient", side_effect=lambda **kw: real_client(transport=httpx.MockTransport(handler), **kw)),
        ):
            with TestClient(app) as client:
                client.cookies.set(identity.COOKIE, self.make_session(email))
                return client.post("/api/v1/me/model-credential")

    def test_model_credential_returns_key_without_persisting_it(self):
        key_value = "sk-this-must-not-be-stored"

        def endpoint(request):
            self.assertEqual(request.url, "https://aethernettech.dev/api/v1/admin/provision/user-key")
            self.assertEqual(request.headers["x-api-key"], "adm_sc_test_restricted-provision-key")
            self.assertEqual(request.headers["user-agent"], sub2api.USER_AGENT)
            self.assertNotIn("authorization", request.headers)
            self.assertEqual(json.loads(request.content), {
                "email": "worker@example.com",
                "display_name": "经办员工",
                "key_name": "qiongyu-office",
                "default_group_id": 7,
                "rate_limit_5h": 20,
            })
            return httpx.Response(200, json={
                "user_id": 51, "user_created": True, "key_id": 82,
                "key_created": True, "key": key_value, "group_id": 7,
                "subscription_id": 9001, "subscription_created": True,
            })

        response = self.call_model_credential("worker@example.com", endpoint)
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json(), {
            "base_url": "https://aethernettech.dev/v1",
            "api_key": key_value,
            "key_id": 82,
            "group_id": 7,
        })
        with store.database() as db:
            audit_rows = db.all("SELECT payload FROM events WHERE type='model.credential_issued'")
            self.assertTrue(audit_rows)
            self.assertTrue(all(key_value not in row["payload"] for row in audit_rows))
            self.assertTrue(any('"key_id":82' in row["payload"] and '"sub2api_user_id":51' in row["payload"]
                and '"subscription_id":9001' in row["payload"] and '"subscription_created":true' in row["payload"]
                for row in audit_rows))
            tables = [row["name"] for row in db.all("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]
            for table in tables:
                for row in db.all(f'SELECT * FROM "{table}"'):
                    self.assertNotIn(key_value, repr(tuple(row)), f"credential persisted in {table}")

    def test_model_credential_accepts_legacy_response_without_subscription_metadata(self):
        response = self.call_model_credential("worker@example.com", lambda request: httpx.Response(200, json={
            "user_id": 51, "user_created": False, "key_id": 82,
            "key_created": False, "key": "sk-legacy", "group_id": 7,
        }))
        self.assertEqual(response.status_code, 200, response.text)
        with store.database() as db:
            audit_rows = db.all("SELECT payload FROM events WHERE type='model.credential_issued'")
        self.assertTrue(any('"subscription_id":null' in row["payload"]
            and '"subscription_created":null' in row["payload"] for row in audit_rows))

    def test_model_credential_omits_unconfigured_optional_rate_limits(self):
        def endpoint(request):
            body = json.loads(request.content)
            self.assertEqual(body["rate_limit_5h"], 20)
            self.assertNotIn("rate_limit_1d", body)
            self.assertNotIn("rate_limit_7d", body)
            return httpx.Response(200, json={
                "user_id": 51, "user_created": False, "key_id": 82,
                "key_created": False, "key": "sk-test", "group_id": 7,
            })

        response = self.call_model_credential("worker@example.com", endpoint)
        self.assertEqual(response.status_code, 200, response.text)

    def test_model_credential_includes_configured_optional_rate_limits(self):
        def endpoint(request):
            body = json.loads(request.content)
            self.assertEqual(body["rate_limit_5h"], 20)
            self.assertEqual(body["rate_limit_1d"], 45.5)
            self.assertEqual(body["rate_limit_7d"], 100)
            return httpx.Response(200, json={
                "user_id": 51, "user_created": False, "key_id": 82,
                "key_created": False, "key": "sk-test", "group_id": 7,
            })

        response = self.call_model_credential("worker@example.com", endpoint, {
            "QO_SUB2API_RATE_LIMIT_1D_USD": "45.5",
            "QO_SUB2API_RATE_LIMIT_7D_USD": "100",
        })
        self.assertEqual(response.status_code, 200, response.text)

    def test_model_credential_preserves_small_finite_decimal_as_json_number(self):
        def endpoint(request):
            self.assertIn(b'"rate_limit_5h":1E-1000', request.content)
            return httpx.Response(200, json={
                "user_id": 51, "user_created": False, "key_id": 82,
                "key_created": False, "key": "sk-test", "group_id": 7,
            })

        response = self.call_model_credential("worker@example.com", endpoint, {
            "QO_SUB2API_RATE_LIMIT_5H_USD": "1e-1000",
        })
        self.assertEqual(response.status_code, 200, response.text)

    def test_model_credential_rejects_invalid_subscription_metadata(self):
        base = {
            "user_id": 51, "user_created": False, "key_id": 82,
            "key_created": False, "key": "sk-test", "group_id": 7,
        }
        for overrides in (
            {"subscription_id": True},
            {"subscription_id": "9001"},
            {"subscription_created": 1},
        ):
            response = self.call_model_credential("worker@example.com", lambda request, values=overrides:
                httpx.Response(200, json={**base, **values}))
            self.assertEqual(response.status_code, 502, response.text)
            self.assertEqual(response.json()["detail"], "模型服务返回内容无效，请联系管理员")

    def test_model_credential_maps_sub2api_errors_and_denies_reviewer(self):
        cases = (
            ("group_not_allowed", 409, 403, "默认模型分组尚未授权给此账号，请联系管理员"),
            ("user_disabled", 403, 403, "模型服务账号已停用，请联系管理员"),
            ("key_disabled", 403, 403, "模型服务密钥已被停用，请联系管理员"),
            ("admin_user_not_provisionable", 403, 403, "管理员账号不能开通员工模型服务"),
            ("group_unavailable", 422, 502, "默认模型分组暂不可用，请联系管理员"),
        )
        for reason, upstream_status, expected_status, expected_detail in cases:
            # sub2api's real error body: numeric code + string reason
            for body in ({"code": upstream_status, "message": "x", "reason": reason}, {"code": reason}):
                response = self.call_model_credential(
                    "worker@example.com",
                    lambda request, value=body, status=upstream_status: httpx.Response(status, json=value),
                )
                self.assertEqual(response.status_code, expected_status, response.text)
                self.assertEqual(response.json()["detail"], expected_detail)
        response = self.call_model_credential(
            "reviewer@example.com",
            lambda request: self.fail("reviewer must be rejected before sub2api is called"),
        )
        self.assertEqual(response.status_code, 403, response.text)

    def test_demo_credential_route_is_not_available(self):
        with patch.object(sub2api, "DEMO", True), TestClient(app) as client:
            response = client.post("/api/v1/me/model-credential")
        self.assertEqual(response.status_code, 404)

    def test_sub2api_enterprise_configuration_validation(self):
        good = {
            "QO_SUB2API_URL": "https://aethernettech.dev",
            "QO_SUB2API_PROVISION_KEY": "adm_sc_test_0123456789abcdef",
            "QO_SUB2API_DEFAULT_GROUP_ID": "7",
            "QO_SUB2API_KEY_NAME": "qiongyu-office",
            "QO_SUB2API_RATE_LIMIT_5H_USD": "20",
        }
        with patch.object(sub2api, "DEMO", False), patch.dict(os.environ, good, clear=False):
            sub2api.validate_config()
            for overrides in (
                {"QO_SUB2API_URL": "http://aethernettech.dev"},
                {"QO_SUB2API_URL": "https://aethernettech.dev/v1"},
                {"QO_SUB2API_DEFAULT_GROUP_ID": "0"},
                {"QO_SUB2API_PROVISION_KEY": ""},
                {"QO_SUB2API_PROVISION_KEY": "   "},
                {"QO_SUB2API_PROVISION_KEY": "restricted-key"},
                {"QO_SUB2API_PROVISION_KEY": "adm_sc_"},
                {"QO_SUB2API_KEY_NAME": "Invalid Name"},
            ):
                with patch.dict(os.environ, overrides, clear=False), self.assertRaises(RuntimeError):
                    sub2api.validate_config()
            with patch.dict(os.environ, {"QO_SUB2API_PROVISION_KEY": "  adm_sc_test_key  "}, clear=False):
                sub2api.validate_config()
        missing_limit = {key: value for key, value in good.items() if key != "QO_SUB2API_RATE_LIMIT_5H_USD"}
        with patch.object(sub2api, "DEMO", False), patch.dict(os.environ, missing_limit, clear=True):
            with self.assertRaisesRegex(RuntimeError, "QO_SUB2API_RATE_LIMIT_5H_USD"):
                sub2api.validate_config()
        for name, values in (
            ("QO_SUB2API_RATE_LIMIT_5H_USD", ("-1", "nan", "inf", "100000.1", "not-a-number", "10_0")),
            ("QO_SUB2API_RATE_LIMIT_1D_USD", ("-1", "nan", "inf", "100000.1", "not-a-number", "10_0")),
            ("QO_SUB2API_RATE_LIMIT_7D_USD", ("-1", "nan", "inf", "100000.1", "not-a-number", "10_0")),
        ):
            for value in values:
                with patch.object(sub2api, "DEMO", False), patch.dict(os.environ, {**good, name: value}, clear=True):
                    with self.assertRaisesRegex(RuntimeError, name):
                        sub2api.validate_config()

    def test_missing_sub2api_5h_limit_fails_app_startup(self):
        startup_env = {
            "QO_SUB2API_URL": "https://aethernettech.dev",
            "QO_SUB2API_PROVISION_KEY": "adm_sc_test_0123456789abcdef",
            "QO_SUB2API_DEFAULT_GROUP_ID": "7",
            "QO_SUB2API_KEY_NAME": "qiongyu-office",
        }
        with (patch.object(sub2api, "DEMO", False), patch.dict(os.environ, startup_env, clear=True)):
            with self.assertRaisesRegex(RuntimeError, "QO_SUB2API_RATE_LIMIT_5H_USD"):
                with TestClient(app):
                    self.fail("enterprise application must reject startup without the 5h limit")

    def test_sqlite_legacy_users_and_flows_get_new_columns(self):
        legacy_dir = Path(self.temp.name) / "legacy"
        legacy_dir.mkdir(exist_ok=True)
        old_data = store.DATA
        store.DATA = legacy_dir
        try:
            raw = sqlite3.connect(legacy_dir / "office.db")
            try:
                raw.execute("CREATE TABLE users (id TEXT PRIMARY KEY, name TEXT NOT NULL, department TEXT NOT NULL, enabled INTEGER NOT NULL DEFAULT 1, role TEXT NOT NULL)")
                raw.execute("CREATE TABLE auth_flows (state TEXT PRIMARY KEY, verifier TEXT NOT NULL, nonce TEXT NOT NULL, desktop TEXT NOT NULL, browser TEXT NOT NULL, expires TEXT NOT NULL)")
                raw.execute("INSERT INTO users(id,name,department,role) VALUES('legacy','Legacy','Ops','employee')")
                raw.commit()
            finally:
                raw.close()
            store.initialize()
            with store.database() as db:
                user_columns = {row["name"] for row in db.all("PRAGMA table_info(users)")}
                flow_columns = {row["name"] for row in db.all("PRAGMA table_info(auth_flows)")}
                self.assertTrue({"email", "subject", "source", "roster_blocked"} <= user_columns)
                self.assertIn("port", flow_columns)
                self.assertEqual(db.one("SELECT name FROM users WHERE id='legacy'")["name"], "Legacy")
        finally:
            store.DATA = old_data


    def test_sub2api_same_origin_proxy_returns_office_base_url_and_streams(self):
        real_client = httpx.AsyncClient
        seen = []

        def handler(request: httpx.Request):
            seen.append((request.method, str(request.url), dict(request.headers), request.content))
            if request.url.path == "/api/v1/admin/provision/user-key":
                return httpx.Response(200, json={
                    "user_id": 61, "user_created": True, "key_id": 91,
                    "key_created": True, "key": "sk-emp-proxy-test", "group_id": 22,
                    "subscription_id": 6, "subscription_created": True,
                })
            if request.url.path == "/v1/models":
                return httpx.Response(200, json={"data": [{"id": "glm-5.3"}]})
            if request.url.path == "/v1/usage":
                return httpx.Response(200, json={"mode": "quota_limited", "rate_limits": []})
            if request.url.path == "/v1/chat/completions":
                return httpx.Response(
                    200,
                    headers={"Content-Type": "text/event-stream"},
                    content=b"data: {\"id\":\"1\"}\n\ndata: [DONE]\n\n",
                )
            return httpx.Response(404, json={"detail": "unexpected"})

        env = {
            "QO_PUBLIC_ORIGIN": "https://office.aethernettech.dev",
            "QO_SUB2API_URL": "https://aethernettech.dev",
            "QO_SUB2API_PROXY_UPSTREAM": "http://127.0.0.1:8081",
            "QO_SUB2API_PROVISION_KEY": "adm_sc_test_restricted-provision-key",
            "QO_SUB2API_DEFAULT_GROUP_ID": "22",
            "QO_SUB2API_KEY_NAME": "qiongyu-office",
            "QO_SUB2API_RATE_LIMIT_5H_USD": "20",
        }
        with (
            patch.object(sub2api, "DEMO", False),
            patch.dict(os.environ, env),
            patch.object(sub2api.httpx, "AsyncClient", side_effect=lambda **kw: real_client(transport=httpx.MockTransport(handler), **kw)),
        ):
            sub2api.validate_config()
            with TestClient(app) as client:
                client.cookies.set(identity.COOKIE, self.make_session("worker@example.com"))
                cred = client.post("/api/v1/me/model-credential")
                self.assertEqual(cred.status_code, 200, cred.text)
                self.assertEqual(cred.json()["base_url"], "https://office.aethernettech.dev/v1")

                # Missing bearer rejected
                self.assertEqual(client.get("/v1/models").status_code, 401)
                # Admin provision key rejected
                self.assertEqual(client.get("/v1/models", headers={"Authorization": "Bearer adm_sc_secret"}).status_code, 403)

                # Valid employee key proxies to loopback upstream without leaking session cookie,
                # and tolerates raw UTF-8 bytes decoded as latin-1 in User-Agent (Electron "穹驭办公/0.2.22")
                models_resp = client.get(
                    "/v1/models",
                    headers=[
                        (b"authorization", b"Bearer sk-emp-proxy-test"),
                        (b"user-agent", "Mozilla/5.0 穹驭办公/0.2.22 Electron/45.3.4".encode("utf-8")),
                    ],
                )
                self.assertEqual(models_resp.status_code, 200)
                self.assertEqual(models_resp.json(), {"data": [{"id": "glm-5.3"}]})
                self.assertEqual(models_resp.headers.get("x-accel-buffering"), "no")

                usage_resp = client.get("/v1/usage", headers={"Authorization": "Bearer sk-emp-proxy-test"})
                self.assertEqual(usage_resp.status_code, 200)
                self.assertEqual(usage_resp.json()["mode"], "quota_limited")

                sse_resp = client.post(
                    "/v1/chat/completions",
                    headers={"Authorization": "Bearer sk-emp-proxy-test"},
                    json={"model": "glm-5.3", "stream": True, "messages": [{"role": "user", "content": "hi"}]},
                )
                self.assertEqual(sse_resp.status_code, 200)
                self.assertIn("text/event-stream", sse_resp.headers.get("content-type", ""))
                self.assertIn("data: [DONE]", sse_resp.text)

        proxy_calls = [item for item in seen if item[1].startswith("http://127.0.0.1:8081/v1/")]
        self.assertEqual(len(proxy_calls), 3)
        for _, _, hdrs, _ in proxy_calls:
            self.assertNotIn("cookie", hdrs)
            self.assertEqual(hdrs.get("authorization"), "Bearer sk-emp-proxy-test")


if __name__ == "__main__":
    unittest.main()
