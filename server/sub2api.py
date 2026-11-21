import json
import os
import re
from decimal import Decimal, InvalidOperation
from urllib.parse import urlsplit

import httpx
from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse, StreamingResponse

from . import identity
from .identity import user_for
from .store import DEMO, audit, database

router = APIRouter()
USER_AGENT = "QiongyuOffice-Server/0.1.0"
PROXY_FORWARD_REQ_HEADERS = (
    "authorization",
    "content-type",
    "accept",
    "openai-beta",
    "x-request-id",
    "idempotency-key",
)
PROXY_FORWARD_RESP_HEADERS = (
    "content-type",
    "retry-after",
    "x-request-id",
    "x-ratelimit-limit-requests",
    "x-ratelimit-limit-tokens",
    "x-ratelimit-remaining-requests",
    "x-ratelimit-remaining-tokens",
    "x-ratelimit-reset-requests",
    "x-ratelimit-reset-tokens",
)
PROXY_TIMEOUT = httpx.Timeout(connect=15.0, read=600.0, write=60.0, pool=15.0)


def settings():
    return {
        "url": os.getenv("QO_SUB2API_URL", "").strip().rstrip("/"),
        "provision_key": os.getenv("QO_SUB2API_PROVISION_KEY", "").strip(),
        "default_group_id": os.getenv("QO_SUB2API_DEFAULT_GROUP_ID", "").strip(),
        "key_name": os.getenv("QO_SUB2API_KEY_NAME", "qiongyu-office").strip(),
        "rate_limit_5h_usd": os.getenv("QO_SUB2API_RATE_LIMIT_5H_USD"),
        "rate_limit_1d_usd": os.getenv("QO_SUB2API_RATE_LIMIT_1D_USD"),
        "rate_limit_7d_usd": os.getenv("QO_SUB2API_RATE_LIMIT_7D_USD"),
        "proxy_upstream": os.getenv("QO_SUB2API_PROXY_UPSTREAM", "").strip().rstrip("/"),
        "client_url": os.getenv("QO_SUB2API_CLIENT_URL", "").strip().rstrip("/"),
    }


def _valid_origin(raw: str, *, allow_loopback_http: bool = False) -> bool:
    try:
        parsed = urlsplit(raw)
    except ValueError:
        return False
    if not raw or not parsed.hostname or parsed.path not in ("", "/") or parsed.query or parsed.fragment or parsed.username or parsed.password:
        return False
    if parsed.scheme == "https":
        return True
    if allow_loopback_http and parsed.scheme == "http" and parsed.hostname in {"127.0.0.1", "localhost", "::1"}:
        return True
    return False


def client_base_url(config) -> str:
    if config.get("client_url"):
        return f"{config['client_url']}/v1"
    if config.get("proxy_upstream"):
        origin = os.getenv("QO_PUBLIC_ORIGIN", identity.ORIGIN).strip().rstrip("/")
        if _valid_origin(origin, allow_loopback_http=False):
            return f"{origin}/v1"
    return f"{config['url']}/v1"


def parse_rate_limit(value, name, *, required=False):
    if value is None or not value.strip():
        if required:
            raise RuntimeError(f"Enterprise mode requires {name}.")
        return None
    value = value.strip()
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", value):
        raise RuntimeError(f"{name} must be a finite number between 0 and 100000.")
    try:
        parsed = Decimal(value)
    except InvalidOperation as error:
        raise RuntimeError(f"{name} must be a finite number between 0 and 100000.") from error
    if not parsed.is_finite() or parsed < 0 or parsed > 100000:
        raise RuntimeError(f"{name} must be a finite number between 0 and 100000.")
    return parsed


def provision_request_json(payload):
    pairs = []
    for key, value in payload.items():
        encoded_value = str(value) if isinstance(value, Decimal) else json.dumps(value)
        pairs.append(f"{json.dumps(key)}:{encoded_value}")
    return "{" + ",".join(pairs) + "}"


def validate_config():
    if DEMO:
        return
    config = settings()
    try:
        url = urlsplit(config["url"])
    except ValueError as error:
        raise RuntimeError("QO_SUB2API_URL must be a valid HTTPS origin.") from error
    try:
        group_id = int(config["default_group_id"])
    except ValueError:
        group_id = 0
    if (not config["url"] or url.scheme != "https" or not url.hostname or url.path not in ("", "/")
            or url.query or url.fragment or url.username or url.password):
        raise RuntimeError("Enterprise mode requires QO_SUB2API_URL to be an HTTPS origin without credentials or a path.")
    provision_key = config["provision_key"]
    if not provision_key.startswith("adm_sc_") or not provision_key[len("adm_sc_"):].strip():
        raise RuntimeError("QO_SUB2API_PROVISION_KEY must use the adm_sc_ restricted-key prefix.")
    if group_id <= 0:
        raise RuntimeError("Enterprise mode requires a positive QO_SUB2API_DEFAULT_GROUP_ID.")
    if not re.fullmatch(r"[a-z0-9][a-z0-9:._-]{0,63}", config["key_name"]):
        raise RuntimeError("QO_SUB2API_KEY_NAME must match the provision key-name format.")
    for key, env_name, required in (
        ("rate_limit_5h_usd", "QO_SUB2API_RATE_LIMIT_5H_USD", True),
        ("rate_limit_1d_usd", "QO_SUB2API_RATE_LIMIT_1D_USD", False),
        ("rate_limit_7d_usd", "QO_SUB2API_RATE_LIMIT_7D_USD", False),
    ):
        parse_rate_limit(config[key], env_name, required=required)
    if config["proxy_upstream"] and not _valid_origin(config["proxy_upstream"], allow_loopback_http=True):
        raise RuntimeError("QO_SUB2API_PROXY_UPSTREAM must be an HTTPS or loopback HTTP origin without a path.")
    if config["client_url"] and not _valid_origin(config["client_url"], allow_loopback_http=False):
        raise RuntimeError("QO_SUB2API_CLIENT_URL must be a valid HTTPS origin without a path.")


ERROR_MESSAGES = {
    "admin_user_not_provisionable": (403, "管理员账号不能开通员工模型服务"),
    "user_disabled": (403, "模型服务账号已停用，请联系管理员"),
    "key_disabled": (403, "模型服务密钥已被停用，请联系管理员"),
    "group_not_allowed": (403, "默认模型分组尚未授权给此账号，请联系管理员"),
    "group_unavailable": (502, "默认模型分组暂不可用，请联系管理员"),
}


def error_for(response):
    # sub2api errors are {"code": <http status int>, "message": str, "reason": "<error id>"};
    # a string "code" is also accepted.
    try:
        body = response.json()
        code = next((value for value in (body.get("reason"), body.get("code")) if isinstance(value, str)), None)
    except (ValueError, AttributeError):
        code = None
    if code in ERROR_MESSAGES:
        return ERROR_MESSAGES[code]
    if response.status_code == 403:
        return 403, "当前账号无权开通模型服务，请联系管理员"
    return 502, "模型服务暂不可用，请稍后重试或联系管理员"


@router.post("/api/v1/me/model-credential")
async def model_credential(request: Request):
    if DEMO:
        raise HTTPException(404, "演示模式不提供模型凭证")
    user = user_for(request)
    if user["role"] not in {"employee", "manager", "admin"}:
        raise HTTPException(403, "当前角色无需模型服务权限")
    config = settings()
    try:
        default_group_id = int(config["default_group_id"])
        rate_limit_5h = parse_rate_limit(
            config["rate_limit_5h_usd"], "QO_SUB2API_RATE_LIMIT_5H_USD", required=True
        )
        rate_limit_1d = parse_rate_limit(config["rate_limit_1d_usd"], "QO_SUB2API_RATE_LIMIT_1D_USD")
        rate_limit_7d = parse_rate_limit(config["rate_limit_7d_usd"], "QO_SUB2API_RATE_LIMIT_7D_USD")
        request_body = {
            "email": user["email"],
            "display_name": user["name"],
            "key_name": config["key_name"],
            "default_group_id": default_group_id,
            "rate_limit_5h": rate_limit_5h,
        }
        if rate_limit_1d is not None:
            request_body["rate_limit_1d"] = rate_limit_1d
        if rate_limit_7d is not None:
            request_body["rate_limit_7d"] = rate_limit_7d
        async with httpx.AsyncClient(timeout=15, trust_env=False) as client:
            response = await client.post(
                f"{config['url']}/api/v1/admin/provision/user-key",
                headers={
                    "x-api-key": config["provision_key"],
                    "User-Agent": USER_AGENT,
                    "Content-Type": "application/json",
                },
                content=provision_request_json(request_body),
            )
    except (httpx.HTTPError, ValueError):
        raise HTTPException(502, "模型服务暂不可用，请稍后重试或联系管理员")
    if response.status_code != 200:
        status, message = error_for(response)
        raise HTTPException(status, message)
    try:
        result = response.json()
        api_key = result["key"]
        sub2api_user_id = result["user_id"]
        key_id = result["key_id"]
        group_id = result["group_id"]
        user_created = result["user_created"]
        key_created = result["key_created"]
        subscription_id = result.get("subscription_id")
        subscription_created = result.get("subscription_created")
        if (not isinstance(api_key, str) or not api_key
                or isinstance(sub2api_user_id, bool) or not isinstance(sub2api_user_id, int)
                or isinstance(key_id, bool) or not isinstance(key_id, int)
                or isinstance(group_id, bool) or not isinstance(group_id, int) or group_id <= 0
                or not isinstance(user_created, bool) or not isinstance(key_created, bool)
                or (subscription_id is not None and (isinstance(subscription_id, bool) or not isinstance(subscription_id, int)))
                or ("subscription_created" in result and not isinstance(subscription_created, bool))):
            raise ValueError("invalid provision response")
    except (ValueError, KeyError, TypeError):
        raise HTTPException(502, "模型服务返回内容无效，请联系管理员")
    with database(write=True) as db:
        audit(
            db,
            "model.credential_issued",
            user["id"],
            sub2api_user_id=sub2api_user_id,
            key_id=key_id,
            user_created=user_created,
            key_created=key_created,
            subscription_id=subscription_id,
            subscription_created=subscription_created,
        )
    return {
        "base_url": client_base_url(config),
        "api_key": api_key,
        "key_id": key_id,
        "group_id": group_id,
    }


def _extract_bearer_token(request: Request) -> str:
    raw = (request.headers.get("authorization") or "").strip()
    if not raw.lower().startswith("bearer "):
        raise HTTPException(401, "缺少模型访问凭证")
    token = raw[7:].strip()
    if not token or len(token) > 512 or any(ch.isspace() for ch in token):
        raise HTTPException(401, "模型访问凭证无效")
    if token.startswith("adm_sc_"):
        raise HTTPException(403, "禁止使用管理密钥调用模型代理")
    return token


async def _proxy_v1(request: Request, endpoint: str):
    if DEMO:
        raise HTTPException(404, "演示模式未开启模型代理")
    config = settings()
    upstream = config["proxy_upstream"]
    if not upstream:
        raise HTTPException(404, "未启用模型代理转发")
    _extract_bearer_token(request)

    target_url = f"{upstream}/v1/{endpoint}"
    if request.url.query:
        target_url = f"{target_url}?{request.url.query}"

    forward_headers = {
        "Accept-Encoding": "identity",
        "User-Agent": USER_AGENT,
    }
    for name in PROXY_FORWARD_REQ_HEADERS:
        val = request.headers.get(name)
        if val:
            safe_val = val.encode("ascii", "ignore").decode("ascii").strip()
            if safe_val:
                forward_headers[name] = safe_val

    body = await request.body() if request.method in {"POST", "PUT", "PATCH"} else None
    client = httpx.AsyncClient(timeout=PROXY_TIMEOUT, trust_env=False)
    try:
        req = client.build_request(
            request.method,
            target_url,
            headers=forward_headers,
            content=body,
        )
        upstream_resp = await client.send(req, stream=True)
    except httpx.HTTPError:
        await client.aclose()
        return JSONResponse(
            {"error": {"message": "模型服务暂不可用，请稍后重试", "type": "upstream_error"}},
            status_code=502,
        )

    resp_headers = {
        "Cache-Control": "no-store",
        "X-Accel-Buffering": "no",
    }
    for name in PROXY_FORWARD_RESP_HEADERS:
        val = upstream_resp.headers.get(name)
        if val is not None:
            resp_headers[name] = val

    async def stream_body():
        try:
            async for chunk in upstream_resp.aiter_bytes():
                if chunk:
                    yield chunk
        finally:
            await upstream_resp.aclose()
            await client.aclose()

    return StreamingResponse(
        stream_body(),
        status_code=upstream_resp.status_code,
        headers=resp_headers,
    )


@router.get("/v1/models")
async def proxy_models(request: Request):
    return await _proxy_v1(request, "models")


@router.get("/v1/usage")
async def proxy_usage(request: Request):
    return await _proxy_v1(request, "usage")


@router.post("/v1/chat/completions")
async def proxy_chat_completions(request: Request):
    return await _proxy_v1(request, "chat/completions")


@router.post("/v1/images/generations")
async def proxy_images_generations(request: Request):
    return await _proxy_v1(request, "images/generations")


@router.post("/v1/responses")
async def proxy_responses(request: Request):
    return await _proxy_v1(request, "responses")
