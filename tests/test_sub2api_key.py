import contextlib
import io
import json
import os
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from unittest.mock import patch

from scripts import check_sub2api_key


class ProbeServer(ThreadingHTTPServer):
    post_status = 400
    users_status = 403
    users_body = b'{"code":"FORBIDDEN","message":"not allowed"}'
    seen = []


class ProbeHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        size = int(self.headers.get("Content-Length", "0"))
        body = json.loads(self.rfile.read(size))
        self.server.seen.append((
            self.command, self.path, self.headers.get("x-api-key"),
            self.headers.get("Authorization"), self.headers.get("User-Agent"), body,
        ))
        self.send_response(self.server.post_status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(b'{"reason":"invalid_email"}')

    def do_GET(self):
        self.server.seen.append((
            self.command, self.path, self.headers.get("x-api-key"),
            self.headers.get("Authorization"), self.headers.get("User-Agent"), None,
        ))
        self.send_response(self.server.users_status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(self.server.users_body)

    def log_message(self, _format, *_args):
        pass


class Sub2apiKeyProbeTest(unittest.TestCase):
    def setUp(self):
        ProbeServer.seen = []
        self.server = ProbeServer(("127.0.0.1", 0), ProbeHandler)
        self.server.users_body = b'{"code":"FORBIDDEN","message":"not allowed"}'
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=2)

    def run_script(self):
        output = io.StringIO()
        with patch.dict(os.environ, {
            "QO_SUB2API_URL": f"http://127.0.0.1:{self.server.server_address[1]}",
            "QO_SUB2API_PROVISION_KEY": "adm_sc_test_secret",
        }, clear=False), contextlib.redirect_stdout(output):
            code = check_sub2api_key.main()
        return code, output.getvalue()

    def test_both_read_only_probes_match_contract(self):
        code, output = self.run_script()
        self.assertEqual(code, 0)
        self.assertIn("PASS", output)
        self.assertEqual([row[1] for row in ProbeServer.seen], [
            "/api/v1/admin/provision/user-key", "/api/v1/admin/users?page=1&page_size=1"
        ])
        self.assertTrue(all(
            row[2] == "adm_sc_test_secret" and row[3] is None
            and row[4] == "QiongyuOffice-DeployCheck/1"
            for row in ProbeServer.seen
        ))
        self.assertEqual(ProbeServer.seen[0][5], {
            "email": "not-an-email", "key_name": "qiongyu-office", "default_group_id": 1
        })
        self.assertNotIn("adm_sc_test_secret", output)

    def test_provision_http_200_is_rejected(self):
        self.server.post_status = 200
        code, output = self.run_script()
        self.assertEqual(code, 1)
        self.assertIn("provision probe", output)
        self.assertNotIn("adm_sc_test_secret", output)

    def test_users_http_200_is_rejected(self):
        self.server.users_status = 200
        code, output = self.run_script()
        self.assertEqual(code, 1)
        self.assertIn("users probe", output)
        self.assertNotIn("adm_sc_test_secret", output)

    def test_users_cloudflare_403_html_is_rejected(self):
        self.server.users_body = b"error code: 1010"
        code, output = self.run_script()
        self.assertEqual(code, 1)
        self.assertIn("users probe: expected HTTP 403 with JSON code FORBIDDEN", output)
        self.assertNotIn("adm_sc_test_secret", output)


if __name__ == "__main__":
    unittest.main()
