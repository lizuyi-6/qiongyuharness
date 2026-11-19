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
        cls.data = Path(cls.temp.name) / "data"
        cls.data.mkdir()
        cls.roster_path = Path(cls.temp.name) / "users.json"
        cls.roster_path.write_text(json.dumps([
            {"email": "Admin@Example.com", "name": "管理员", "department": "运营", "role": "admin"},
            {"email": "worker@example.com", "name": "经办员工", "department": "产品", "role": "employee"},
            {"email": "unverified@example.com", "name": "未验证用户", "department": "产品", "role": "employee"},
            {"email": "reviewer@example.com", "name": "审核员", "department": "运营", "role": "reviewer"},
        ]), encoding="utf-8")
        store.DATA = cls.data
        store.DEMO = False
        os.environ["QO_USERS_FILE"] = str(cls.roster_path)
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
        cls.temp.cleanup()

    def setUp(self):
        store.initialize()

    def user_id(self, email):
        import hashlib
        return "u_" + hashlib.sha256(email.lower().encode("utf-8")).hexdigest()[:24]

    def oidc_login(self, subject, email=None, email_verified=None, *, code="login", desktop=False):
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

    def make_session(self, email):
        with store.database(write=True) as db:
            return identity.make_session(db, self.user_id(email))

    def call_model_credential(self, email, handler):
        real_client = httpx.AsyncClient
        with (
            patch.object(sub2api, "DEMO", False),
            patch.dict(os.environ, {
                "QO_SUB2API_URL": "https://aethernettech.dev",
                "QO_SUB2API_PROVISION_KEY": "restricted-provision-key",
                "QO_SUB2API_DEFAULT_GROUP_ID": "7",
                "QO_SUB2API_KEY_NAME": "qiongyu-office",
            }),
            patch.object(sub2api.httpx, "AsyncClient", side_effect=lambda **kw: real_client(transport=httpx.MockTransport(handler), **kw)),
        ):
            with TestClient(app) as client:
                client.cookies.set(identity.COOKIE, self.make_session(email))
                return client.post("/api/v1/me/model-credential")

    def test_model_credential_returns_key_without_persisting_it(self):
        key_value = "sk-this-must-not-be-stored"

        def endpoint(request):
            self.assertEqual(request.url, "https://aethernettech.dev/api/v1/admin/provision/user-key")
            self.assertEqual(request.headers["authorization"], "Bearer restricted-provision-key")
            self.assertEqual(json.loads(request.content), {
                "email": "worker@example.com",
                "display_name": "经办员工",
                "key_name": "qiongyu-office",
                "default_group_id": 7,
            })
            return httpx.Response(200, json={
                "user_id": 51, "user_created": True, "key_id": 82,
                "key_created": True, "key": key_value, "group_id": 7,
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
            self.assertTrue(any('"key_id":82' in row["payload"] and '"sub2api_user_id":51' in row["payload"] for row in audit_rows))
            tables = [row["name"] for row in db.all("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]
            for table in tables:
                for row in db.all(f'SELECT * FROM "{table}"'):
                    self.assertNotIn(key_value, repr(tuple(row)), f"credential persisted in {table}")

    def test_model_credential_maps_sub2api_errors_and_denies_reviewer(self):
        for code, upstream_status in (("group_not_allowed", 409), ("user_disabled", 403)):
            response = self.call_model_credential(
                "worker@example.com",
                lambda request, value=code, status=upstream_status: httpx.Response(status, json={"code": value}),
            )
            self.assertEqual(response.status_code, 403, response.text)
            self.assertTrue(response.json()["detail"])
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
            "QO_SUB2API_PROVISION_KEY": "restricted-key",
            "QO_SUB2API_DEFAULT_GROUP_ID": "7",
            "QO_SUB2API_KEY_NAME": "qiongyu-office",
        }
        with patch.object(sub2api, "DEMO", False), patch.dict(os.environ, good, clear=False):
            sub2api.validate_config()
            for overrides in (
                {"QO_SUB2API_URL": "http://aethernettech.dev"},
                {"QO_SUB2API_URL": "https://aethernettech.dev/v1"},
                {"QO_SUB2API_DEFAULT_GROUP_ID": "0"},
                {"QO_SUB2API_PROVISION_KEY": ""},
                {"QO_SUB2API_KEY_NAME": "Invalid Name"},
            ):
                with patch.dict(os.environ, overrides, clear=False), self.assertRaises(RuntimeError):
                    sub2api.validate_config()

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
                self.assertTrue({"email", "subject"} <= user_columns)
                self.assertIn("port", flow_columns)
                self.assertEqual(db.one("SELECT name FROM users WHERE id='legacy'")["name"], "Legacy")
        finally:
            store.DATA = old_data


if __name__ == "__main__":
    unittest.main()
