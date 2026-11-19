"""Read-only deployment check for the restricted sub2api provision key."""

import json
import os
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


def _request(url, method, key, body=None):
    headers = {"x-api-key": key, "Accept": "application/json"}
    data = None
    if body is not None:
        data = json.dumps(body, separators=(",", ":")).encode("utf-8")
        headers["Content-Type"] = "application/json"
    request = Request(url, data=data, headers=headers, method=method)
    try:
        return urlopen(request, timeout=10)
    except HTTPError as response:
        return response


def _body(response):
    try:
        return json.loads(response.read().decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError, AttributeError):
        return None


def run_checks(base_url, key, request=_request):
    """Return failed probe labels; never include the credential in diagnostics."""
    origin = base_url.strip().rstrip("/")
    failures = []

    try:
        response = request(
            f"{origin}/api/v1/admin/provision/user-key",
            "POST",
            key,
            {"email": "not-an-email", "key_name": "qiongyu-office", "default_group_id": 1},
        )
        try:
            payload = _body(response)
            reasons = (payload.get("reason"), payload.get("code")) if isinstance(payload, dict) else ()
            if response.status != 400 or "invalid_email" not in reasons:
                failures.append("provision probe: expected HTTP 400 with reason/code invalid_email")
        finally:
            response.close()
    except (OSError, URLError, TimeoutError, ValueError):
        failures.append("provision probe: request failed")

    try:
        response = request(f"{origin}/api/v1/admin/users?page=1&page_size=1", "GET", key)
        try:
            if response.status != 403:
                failures.append("users probe: expected HTTP 403")
        finally:
            response.close()
    except (OSError, URLError, TimeoutError, ValueError):
        failures.append("users probe: request failed")

    return failures


def main():
    base_url = os.getenv("QO_SUB2API_URL", "").strip()
    key = os.getenv("QO_SUB2API_PROVISION_KEY", "").strip()
    if not base_url or not key:
        print("FAIL: set QO_SUB2API_URL and QO_SUB2API_PROVISION_KEY")
        return 2
    failures = run_checks(base_url, key)
    if failures:
        for failure in failures:
            print(f"FAIL: {failure}")
        return 1
    print("PASS: provision permission and restricted users access probes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
