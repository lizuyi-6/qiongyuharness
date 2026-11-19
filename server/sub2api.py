import os
import re
from urllib.parse import urlsplit

import httpx
from fastapi import APIRouter, HTTPException, Request

from .identity import user_for
from .store import DEMO, audit, database

router = APIRouter()
USER_AGENT = "QiongyuOffice-Server/0.1.0"


def settings():
    return {
        "url": os.getenv("QO_SUB2API_URL", "").strip().rstrip("/"),
        "provision_key": os.getenv("QO_SUB2API_PROVISION_KEY", "").strip(),
        "default_group_id": os.getenv("QO_SUB2API_DEFAULT_GROUP_ID", "").strip(),
        "key_name": os.getenv("QO_SUB2API_KEY_NAME", "qiongyu-office").strip(),
    }


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
        async with httpx.AsyncClient(timeout=15, trust_env=False) as client:
            response = await client.post(
                f"{config['url']}/api/v1/admin/provision/user-key",
                headers={"x-api-key": config["provision_key"], "User-Agent": USER_AGENT},
                json={
                    "email": user["email"],
                    "display_name": user["name"],
                    "key_name": config["key_name"],
                    "default_group_id": default_group_id,
                },
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
        if (not isinstance(api_key, str) or not api_key
                or isinstance(sub2api_user_id, bool) or not isinstance(sub2api_user_id, int)
                or isinstance(key_id, bool) or not isinstance(key_id, int)
                or isinstance(group_id, bool) or not isinstance(group_id, int) or group_id <= 0
                or not isinstance(user_created, bool) or not isinstance(key_created, bool)):
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
        )
    return {
        "base_url": f"{config['url']}/v1",
        "api_key": api_key,
        "key_id": key_id,
        "group_id": group_id,
    }
