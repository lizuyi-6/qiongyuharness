import base64
import hashlib
import os
import re
import secrets
from datetime import UTC, datetime, timedelta
from urllib.parse import urlencode, urlsplit

import httpx
import jwt
from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from pydantic import BaseModel, Field

from .store import DEMO, audit, database, instant

router = APIRouter()
COOKIE = "qo_session"
FLOW_COOKIE = "qo_browser"
ISSUER = os.getenv("QO_OIDC_ISSUER", "").strip()
CLIENT = os.getenv("QO_OIDC_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("QO_OIDC_CLIENT_SECRET", "")
ORIGIN = os.getenv("QO_PUBLIC_ORIGIN", "http://127.0.0.1:8765").rstrip("/")
SECURE = ORIGIN.startswith("https://")
SESSION_HOURS = 8
SAFE_TOKEN = re.compile(r"[A-Za-z0-9_-]{43,128}")


def fail(code, message):
    raise HTTPException(code, message)


def digest(value):
    return hashlib.sha256(value.encode()).hexdigest()


def challenge(value):
    return base64.urlsafe_b64encode(hashlib.sha256(value.encode()).digest()).decode().rstrip("=")


def make_session(db, user_id):
    token = secrets.token_urlsafe(32)
    db.execute("DELETE FROM sessions WHERE expires<?", (instant(),))
    expires = instant(datetime.now(UTC) + timedelta(hours=SESSION_HOURS))
    db.execute("INSERT INTO sessions(token,user_id,expires) VALUES(?,?,?)", (digest(token), user_id, expires))
    return token


def session_cookie(response, token):
    response.set_cookie(COOKIE, token, httponly=True, secure=SECURE, samesite="lax", path="/")


def user_for(request: Request):
    token = request.cookies.get(COOKIE, "")
    if not SAFE_TOKEN.fullmatch(token):
        fail(401, "请先登录")
    with database() as db:
        user = db.one(
            "SELECT u.id,u.name,u.department,u.role FROM users u JOIN sessions s ON u.id=s.user_id WHERE s.token=? AND s.expires>? AND u.enabled=1",
            (digest(token), instant()),
        )
    if not user:
        fail(401, "登录已过期，请重新登录")
    return dict(user)


def validate_identity_config():
    if DEMO:
        return
    origin = urlsplit(ORIGIN)
    if (not SECURE or not origin.hostname or origin.path or origin.query or origin.fragment
            or origin.username or not CLIENT or not ISSUER.startswith("https://")):
        raise RuntimeError("Enterprise mode requires HTTPS QO_PUBLIC_ORIGIN, QO_OIDC_ISSUER and QO_OIDC_CLIENT_ID.")


async def metadata(client):
    response = await client.get(f"{ISSUER.rstrip('/')}/.well-known/openid-configuration")
    response.raise_for_status()
    value = response.json()
    if value.get("issuer") != ISSUER:
        raise ValueError("issuer mismatch")
    for key in ("authorization_endpoint", "token_endpoint", "jwks_uri"):
        url = urlsplit(value[key])
        if url.scheme != "https" or not url.hostname or url.username or url.password or url.fragment:
            raise ValueError("invalid OIDC endpoint")
    return value


@router.get("/auth/start")
async def start_auth(request: Request, desktop: str = ""):
    if not ISSUER or not CLIENT:
        fail(503, "企业登录配置尚未启用")
    if desktop and not SAFE_TOKEN.fullmatch(desktop):
        fail(400, "桌面登录请求无效")
    try:
        async with httpx.AsyncClient(timeout=10, trust_env=False) as client:
            info = await metadata(client)
    except (httpx.HTTPError, ValueError, KeyError, TypeError):
        fail(502, "无法连接企业身份服务")
    state = secrets.token_urlsafe(32)
    nonce = secrets.token_urlsafe(32)
    verifier = secrets.token_urlsafe(48)
    browser = secrets.token_urlsafe(32)
    expires = instant(datetime.now(UTC) + timedelta(minutes=8))
    with database(write=True) as db:
        db.execute("DELETE FROM auth_flows WHERE expires<?", (instant(),))
        db.execute("DELETE FROM desktop_grants WHERE expires<?", (instant(),))
        db.execute(
            "INSERT INTO auth_flows(state,verifier,nonce,desktop,browser,expires) VALUES(?,?,?,?,?,?)",
            (digest(state), verifier, nonce, desktop, digest(browser), expires),
        )
    params = urlencode({
        "client_id": CLIENT, "redirect_uri": f"{ORIGIN}/auth/callback", "response_type": "code",
        "scope": "openid profile", "state": state, "nonce": nonce, "code_challenge_method": "S256",
        "code_challenge": challenge(verifier),
    })
    separator = "&" if "?" in info["authorization_endpoint"] else "?"
    response = RedirectResponse(info["authorization_endpoint"] + separator + params, status_code=303)
    response.set_cookie(FLOW_COOKIE, browser, max_age=480, httponly=True, secure=SECURE, samesite="lax", path="/auth")
    return response


@router.get("/auth/callback")
async def auth_callback(request: Request, state: str = "", code: str = ""):
    if not SAFE_TOKEN.fullmatch(state) or not code or len(code) > 4096:
        fail(400, "企业登录返回内容无效")
    with database(write=True) as db:
        flow = db.one("SELECT * FROM auth_flows WHERE state=? AND browser=? AND expires>?",
                      (digest(state), digest(request.cookies.get(FLOW_COOKIE, "")), instant()))
        if not flow:
            fail(400, "登录请求已失效，请重新发起")
        db.execute("DELETE FROM auth_flows WHERE state=?", (digest(state),))
    try:
        async with httpx.AsyncClient(timeout=12, trust_env=False) as client:
            info = await metadata(client)
            values = {"client_id": CLIENT, "grant_type": "authorization_code", "code": code,
                      "code_verifier": flow["verifier"], "redirect_uri": f"{ORIGIN}/auth/callback"}
            methods = info.get("token_endpoint_auth_methods_supported", ["client_secret_basic"])
            auth = None
            if CLIENT_SECRET:
                if "client_secret_basic" in methods:
                    auth = (CLIENT, CLIENT_SECRET)
                elif "client_secret_post" in methods:
                    values["client_secret"] = CLIENT_SECRET
                else:
                    raise ValueError("unsupported token endpoint authentication")
            exchange = await client.post(info["token_endpoint"], data=values, auth=auth)
            exchange.raise_for_status()
            token = exchange.json()["id_token"]
            header = jwt.get_unverified_header(token)
            keys = await client.get(info["jwks_uri"])
            keys.raise_for_status()
            valid_algs = ("RS256", "RS384", "RS512", "ES256", "ES384", "ES512", "PS256", "PS384", "PS512")
            if header.get("alg") not in valid_algs or not header.get("kid"):
                raise ValueError("unsupported signing key")
            key = next(k for k in keys.json()["keys"] if k.get("kid") == header["kid"] and k.get("use", "sig") == "sig")
            claims = jwt.decode(token, jwt.PyJWK.from_dict(key, algorithm=header["alg"]).key,
                                algorithms=[header["alg"]], audience=CLIENT, issuer=ISSUER, leeway=15,
                                options={"require": ["iss", "sub", "aud", "exp", "iat", "nonce"]})
            if claims["nonce"] != flow["nonce"]:
                raise ValueError("nonce mismatch")
            audience = claims["aud"]
            azp = claims.get("azp")
            if (azp is not None or isinstance(audience, list) and len(audience) > 1) and azp != CLIENT:
                raise ValueError("authorized party mismatch")
    except (httpx.HTTPError, jwt.PyJWTError, ValueError, KeyError, TypeError, StopIteration):
        fail(401, "企业身份验证失败，请重新登录")
    with database(write=True) as db:
        user = db.one("SELECT id FROM users WHERE id=? AND enabled=1", (claims["sub"],))
        if not user:
            fail(403, "此企业账号尚未开通工作台权限")
        audit(db, "auth.login", user["id"], mode="oidc")
        if flow["desktop"]:
            expires = instant(datetime.now(UTC) + timedelta(minutes=2))
            db.execute("INSERT INTO desktop_grants(challenge,user_id,expires) VALUES(?,?,?) "
                       "ON CONFLICT(challenge) DO UPDATE SET user_id=excluded.user_id,expires=excluded.expires",
                       (flow["desktop"], user["id"], expires))
            response = HTMLResponse(
                '<html lang="zh-CN"><meta charset="utf-8"><title>登录完成</title>'
                '<body><h1>企业登录完成</h1><p>请返回穹驭办公桌面应用。</p></body></html>'
            )
        else:
            response = RedirectResponse("/", status_code=303)
            session_cookie(response, make_session(db, user["id"]))
    response.delete_cookie(FLOW_COOKIE, path="/auth", secure=SECURE, httponly=True)
    return response


class DesktopGrant(BaseModel):
    verifier: str = Field(min_length=43, max_length=128, pattern=r"^[A-Za-z0-9_-]+$")


@router.post("/api/v1/auth/desktop/claim")
def desktop_claim(body: DesktopGrant):
    with database(write=True) as db:
        key = challenge(body.verifier)
        row = db.one("SELECT g.user_id FROM desktop_grants g JOIN users u ON u.id=g.user_id WHERE g.challenge=? AND g.expires>? AND u.enabled=1", (key, instant()))
        if not row:
            fail(404, "等待企业登录")
        db.execute("DELETE FROM desktop_grants WHERE challenge=?", (key,))
        return {"session": make_session(db, row["user_id"])}


class DemoLogin(BaseModel):
    user_id: str = Field(min_length=1, max_length=64)


@router.post("/api/v1/auth/demo")
def demo_auth(body: DemoLogin):
    from fastapi.responses import JSONResponse
    if not DEMO or body.user_id not in {"admin", "manager", "lin", "chen", "song", "reviewer"}:
        fail(404, "账号不存在")
    with database(write=True) as db:
        row = db.one("SELECT id,name,department,role FROM users WHERE id=? AND enabled=1", (body.user_id,))
        if not row:
            fail(404, "账号不存在")
        response = JSONResponse(dict(row))
        session_cookie(response, make_session(db, body.user_id))
    return response


@router.post("/api/v1/auth/logout")
def logout(request: Request):
    from fastapi.responses import JSONResponse
    token = request.cookies.get(COOKIE, "")
    if SAFE_TOKEN.fullmatch(token):
        with database(write=True) as db:
            db.execute("DELETE FROM sessions WHERE token=?", (digest(token),))
    response = JSONResponse({"ok": True})
    response.delete_cookie(COOKIE, path="/", secure=SECURE, httponly=True)
    return response
