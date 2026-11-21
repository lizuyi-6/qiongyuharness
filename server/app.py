import asyncio
import logging
import os
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import Depends, FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.trustedhost import TrustedHostMiddleware

from . import conversations, feedback, files, gateway, identity, pipeline, presentation, projects, search, sub2api, tasks, updates
from .store import DEMO, ZONE, initialize

STATIC = Path(__file__).parent / "static"
LOG = logging.getLogger("qiongyu")
MAX_BODY = 14 * 1024 * 1024


class RequestLimits:
    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        headers = dict(scope["headers"])
        method = scope["method"]
        count = 0
        if method in {"POST", "PUT", "PATCH", "DELETE"}:
            origin = headers.get(b"origin")
            site = headers.get(b"sec-fetch-site")
            if (origin and origin.decode("latin1").rstrip("/") != identity.ORIGIN) or site == b"cross-site":
                await JSONResponse({"detail": "请求来源与工作台地址不符"}, status_code=403)(scope, receive, send)
                return
            if headers.get(b"content-length", b"0").isdigit() and int(headers.get(b"content-length", b"0")) > MAX_BODY:
                await JSONResponse({"detail": "上传文件超过大小限制"}, status_code=413)(scope, receive, send)
                return
        async def bounded():
            nonlocal count
            message = await receive()
            count += len(message.get("body", b""))
            if count > MAX_BODY:
                from fastapi import HTTPException
                raise HTTPException(status_code=413, detail="上传文件超过大小限制")
            return message

        async def secured(message):
            if message["type"] == "http.response.start":
                response_headers = message.setdefault("headers", [])
                response_headers.extend([
                    (b"content-security-policy", b"default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"),
                    (b"x-content-type-options", b"nosniff"),
                    (b"referrer-policy", b"no-referrer"),
                    (b"permissions-policy", b"camera=(), microphone=(), geolocation=()"),
                ])
                if not any(name.lower() == b"cache-control" for name, _ in response_headers):
                    response_headers.append((b"cache-control", b"no-store"))
                if identity.SECURE:
                    response_headers.append((b"strict-transport-security", b"max-age=31536000"))
            await send(message)
        await self.app(scope, bounded, secured)


async def scheduler():
    while True:
        try:
            await asyncio.to_thread(tasks.tick)
        except Exception as error:
            LOG.error("Task scheduler error: %s", type(error).__name__)
        await asyncio.sleep(3)


@asynccontextmanager
async def lifespan(_):
    identity.validate_identity_config()
    pipeline.validate_config()
    sub2api.validate_config()
    gateway.validate()
    initialize()
    periodic = asyncio.create_task(scheduler())
    try:
        yield
    finally:
        periodic.cancel()
        try:
            await periodic
        except asyncio.CancelledError:
            pass


app = FastAPI(title="穹驭办公 API", version="0.2.23", lifespan=lifespan,
              docs_url=None, redoc_url=None, openapi_url="/openapi.json" if DEMO else None)
app.add_middleware(RequestLimits)
if not DEMO:
    from urllib.parse import urlsplit
    app.add_middleware(TrustedHostMiddleware, allowed_hosts=[urlsplit(identity.ORIGIN).hostname or "invalid"])
app.include_router(identity.router)
app.include_router(pipeline.router)
app.include_router(tasks.router)
app.include_router(files.router)
app.include_router(conversations.router)
app.include_router(projects.router)
app.include_router(presentation.router)
app.include_router(sub2api.router)
app.include_router(search.router)
app.include_router(updates.router)
app.include_router(feedback.router)
app.mount("/static", StaticFiles(directory=STATIC), name="static")


@app.get("/healthz")
def health():
    return {"status": "ok"}


@app.get("/api/v1/config")
def app_config():
    return {
        "demo": DEMO,
        "sso": bool(identity.ISSUER and identity.CLIENT),
        "timezone": str(ZONE),
        "version": os.environ.get("QO_APP_VERSION") or app.version,
        "web_search": search.search_status(),
    }


@app.get("/api/v1/engine", dependencies=[Depends(identity.user_for)])
def status():
    return gateway.info()


@app.get("/", include_in_schema=False)
def index():
    return FileResponse(STATIC / "index.html")
