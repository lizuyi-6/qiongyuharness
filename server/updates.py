"""桌面客户端增量/全量更新发布与分发接口。"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, Query
from fastapi.responses import FileResponse

from .identity import fail, user_for
from .store import DATA

router = APIRouter(prefix="/api/v1/updates")

RELEASES_DIR = Path(os.environ.get("QO_RELEASES_DIR") or (DATA.parent / "releases"))
SAFE_FILE_RE = re.compile(r"^[A-Za-z0-9._-]+\.(?:qopatch|exe)$")
SEMVER_RE = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)")


def parse_version(value: str | None) -> tuple[int, int, int] | None:
    if not value:
        return None
    match = SEMVER_RE.match(str(value).strip())
    if not match:
        return None
    return int(match.group(1)), int(match.group(2)), int(match.group(3))


def is_newer(latest: str | None, current: str | None) -> bool:
    latest_v = parse_version(latest)
    if not latest_v:
        return False
    current_v = parse_version(current)
    if not current_v:
        return True
    return latest_v > current_v


def load_manifest() -> dict[str, Any] | None:
    manifest_path = RELEASES_DIR / "manifest.json"
    if not manifest_path.is_file():
        return None
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict) or not isinstance(data.get("version"), str):
        return None
    return data


def _artifact_info(entry: Any) -> dict[str, Any] | None:
    if not isinstance(entry, dict):
        return None
    filename = str(entry.get("file") or "").strip()
    sha256 = str(entry.get("sha256") or "").strip().lower()
    if not filename or not SAFE_FILE_RE.match(filename):
        return None
    target = RELEASES_DIR / filename
    if not target.is_file():
        return None
    size = target.stat().st_size
    if size <= 0:
        return None
    return {
        "file": filename,
        "size": size,
        "sha256": sha256,
        "url": f"/api/v1/updates/files/{filename}",
    }


def _can_use_delta(manifest: dict[str, Any], current: str | None, runtime_hash: str | None, delta: dict[str, Any] | None) -> bool:
    if not delta:
        return False
    expected_hash = str(manifest.get("runtime_hash") or "").strip().lower()
    client_hash = str(runtime_hash or "").strip().lower()
    if not expected_hash or not client_hash or expected_hash != client_hash:
        return False
    min_delta = manifest.get("min_delta_version")
    if min_delta:
        min_v = parse_version(str(min_delta))
        cur_v = parse_version(current)
        if min_v and (not cur_v or cur_v < min_v):
            return False
    return True


@router.get("/latest", dependencies=[Depends(user_for)])
def latest_update(
    current: str | None = Query(default=None, max_length=64),
    runtime_hash: str | None = Query(default=None, max_length=128),
):
    manifest = load_manifest()
    if not manifest:
        return {
            "available": False,
            "has_update": False,
            "latest_version": "",
            "update_mode": "none",
            "published_at": "",
            "notes": "",
            "runtime_hash": "",
            "delta": None,
            "full": None,
        }

    latest_version = str(manifest.get("version") or "").strip()
    delta = _artifact_info(manifest.get("delta"))
    full = _artifact_info(manifest.get("full"))
    has_any_pkg = bool(delta or full)
    has_update = has_any_pkg and is_newer(latest_version, current)

    update_mode = "none"
    if has_update:
        if _can_use_delta(manifest, current, runtime_hash, delta):
            update_mode = "delta"
        elif full:
            update_mode = "full"
        elif delta:
            update_mode = "delta"
        else:
            has_update = False

    return {
        "available": has_any_pkg,
        "has_update": has_update,
        "latest_version": latest_version,
        "update_mode": update_mode,
        "published_at": str(manifest.get("published_at") or ""),
        "notes": str(manifest.get("notes") or ""),
        "runtime_hash": str(manifest.get("runtime_hash") or ""),
        "delta": delta,
        "full": full,
    }


@router.get("/files/{filename}", dependencies=[Depends(user_for)])
def download_update_file(filename: str):
    if not SAFE_FILE_RE.match(filename):
        fail(404, "更新文件不存在")
    manifest = load_manifest()
    if not manifest:
        fail(404, "暂无可用更新")
    allowed = set()
    for key in ("delta", "full"):
        item = manifest.get(key)
        if isinstance(item, dict) and isinstance(item.get("file"), str):
            allowed.add(item["file"].strip())
    if filename not in allowed:
        fail(404, "更新文件已过期或不存在")
    target = (RELEASES_DIR / filename).resolve()
    releases_root = RELEASES_DIR.resolve()
    if target.parent != releases_root or not target.is_file():
        fail(404, "更新文件不存在")
    media_type = "application/gzip" if filename.endswith(".qopatch") else "application/vnd.microsoft.portable-executable"
    return FileResponse(
        path=str(target),
        media_type=media_type,
        filename=filename,
        headers={"Cache-Control": "no-store"},
    )
