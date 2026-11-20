"""DeepSeek 官方 Search & Fetch API 服务端代理（集中保管密钥，供桌面 Harness 回环调用）。"""

from __future__ import annotations

import asyncio
import ipaddress
import logging
import os
from typing import Any
from urllib.parse import urlsplit

import httpx
from fastapi import APIRouter, Depends
from pydantic import BaseModel, ConfigDict, Field

from .identity import fail, user_for
from .store import audit, database

router = APIRouter(prefix="/api/v1/search")
LOG = logging.getLogger("qiongyu.search")

DEFAULT_TOOLS_BASE_URL = "https://api.deepseek.com/v1/tools"
DEFAULT_ANTHROPIC_BASE_URL = "https://api.deepseek.com/anthropic/v1"
DEFAULT_ANTHROPIC_MODEL = "deepseek-v4-flash"
USER_AGENT = "deepseek-harness/0.0.1"
MAX_FETCH_CHARS = 24_000


def _config() -> dict[str, Any]:
    api_key = (
        os.getenv("QO_DEEPSEEK_SEARCH_API_KEY", "").strip()
        or os.getenv("DEEPSEEK_API_KEY", "").strip()
    )
    base_url = (
        os.getenv("QO_DEEPSEEK_SEARCH_BASE_URL", "").strip()
        or DEFAULT_TOOLS_BASE_URL
    ).rstrip("/")
    anthropic_url = (
        os.getenv("QO_DEEPSEEK_SEARCH_ANTHROPIC_URL", "").strip()
        or DEFAULT_ANTHROPIC_BASE_URL
    ).rstrip("/")
    model = (
        os.getenv("QO_DEEPSEEK_SEARCH_MODEL", "").strip()
        or DEFAULT_ANTHROPIC_MODEL
    )
    required_raw = os.getenv("QO_DEEPSEEK_SEARCH_REQUIRED", "1").strip().lower()
    required = required_raw not in {"0", "false", "no", "off"}
    return {
        "api_key": api_key,
        "base_url": base_url,
        "anthropic_url": anthropic_url,
        "model": model,
        "required": required,
    }


def search_status() -> dict[str, Any]:
    cfg = _config()
    configured = bool(cfg["api_key"])
    return {
        "configured": configured,
        "required": bool(cfg["required"]),
        "active": bool(configured),
        "provider": "deepseek-search-api",
    }


class SearchQueryRequest(BaseModel):
    model_config = ConfigDict(extra="ignore")
    queries: list[str] = Field(min_length=1, max_length=4)
    count: int = Field(default=8, ge=1, le=20)
    ori_question: str | None = Field(default=None, max_length=500)
    need_content: bool = False


class FetchUrlRequest(BaseModel):
    model_config = ConfigDict(extra="ignore")
    url: str = Field(min_length=8, max_length=2048)


def _is_public_http_url(raw_url: str) -> bool:
    try:
        parsed = urlsplit(raw_url.strip())
    except ValueError:
        return False
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        return False
    host = parsed.hostname.strip().lower()
    if host in {"localhost", "ip6-localhost", "metadata.google.internal"} or host.endswith(".local"):
        return False
    try:
        ip = ipaddress.ip_address(host)
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
            return False
    except ValueError:
        pass
    return True


def _bounded_question(value: str | None) -> str | None:
    if not value:
        return None
    compact = " ".join(str(value).split()).strip()
    return compact[:128] if compact else None


def _map_tools_search_results(payload: Any) -> list[dict[str, str]]:
    if not isinstance(payload, dict) or not isinstance(payload.get("results"), list):
        raise ValueError("DeepSeek Search API 返回结构无效")
    sources: list[dict[str, str]] = []
    for item in payload["results"]:
        if not isinstance(item, dict):
            continue
        url = str(item.get("url") or "").strip()
        if not url:
            continue
        entry: dict[str, str] = {"url": url}
        title = str(item.get("title") or "").strip()
        snippet = str(item.get("snippet") or "").strip()
        summary = str(item.get("summary") or "").strip()
        date = str(item.get("date") or "").strip()
        if title:
            entry["title"] = title
        if snippet:
            entry["snippet"] = snippet
        if summary:
            entry["summary"] = summary
        if date:
            entry["published_at"] = date
        sources.append(entry)
    return sources


def _map_anthropic_search_results(payload: Any) -> tuple[list[dict[str, str]], str]:
    if not isinstance(payload, dict) or not isinstance(payload.get("content"), list):
        raise ValueError("DeepSeek Messages 搜索返回结构无效")
    blocks = payload["content"]
    snippets: dict[str, str] = {}
    answer_parts: list[str] = []
    for block in blocks:
        if not isinstance(block, dict) or block.get("type") != "text":
            continue
        text = str(block.get("text") or "").strip()
        if text:
            answer_parts.append(text)
        for cite in block.get("citations") or []:
            if not isinstance(cite, dict):
                continue
            url = str(cite.get("url") or "").strip()
            cited = str(cite.get("cited_text") or "").strip()
            if url and cited and url not in snippets:
                snippets[url] = cited

    seen: set[str] = set()
    sources: list[dict[str, str]] = []
    for block in blocks:
        if not isinstance(block, dict) or block.get("type") != "web_search_tool_result":
            continue
        for item in block.get("content") or []:
            if not isinstance(item, dict) or item.get("type") != "web_search_result":
                continue
            url = str(item.get("url") or "").strip()
            if not url or url in seen:
                continue
            seen.add(url)
            entry: dict[str, str] = {"url": url}
            title = str(item.get("title") or "").strip()
            page_age = str(item.get("page_age") or "").strip()
            if title:
                entry["title"] = title
            if url in snippets:
                entry["snippet"] = snippets[url]
            if page_age:
                entry["published_at"] = page_age
            sources.append(entry)
    return sources, "\n".join(answer_parts).strip()


async def _search_single_tools(
    client: httpx.AsyncClient,
    cfg: dict[str, Any],
    query: str,
    count: int,
    ori_question: str | None,
    need_content: bool,
) -> tuple[list[dict[str, str]], str, str]:
    endpoint = f"{cfg['base_url']}/search"
    body: dict[str, Any] = {
        "query": query[:128],
        "count": count,
        "need_content": bool(need_content),
    }
    if ori_question:
        body["ori_question"] = ori_question
    response = await client.post(
        endpoint,
        headers={
            "Authorization": f"Bearer {cfg['api_key']}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": USER_AGENT,
        },
        json=body,
    )
    if response.status_code in {404, 405}:
        return await _search_single_anthropic(client, cfg, query)
    if response.status_code == 401:
        fail(502, "DeepSeek 搜索密钥无效或已过期，请检查 QO_DEEPSEEK_SEARCH_API_KEY")
    if response.status_code == 402:
        fail(502, "DeepSeek 搜索账户余额不足，请充值后重试")
    if response.status_code != 200:
        detail = ""
        try:
            err = response.json()
            if isinstance(err, dict):
                raw = err.get("error")
                detail = raw.get("message", "") if isinstance(raw, dict) else str(raw or err.get("message") or "")
        except ValueError:
            pass
        LOG.warning("DeepSeek search failed HTTP %s: %s", response.status_code, detail[:200])
        fail(502, f"DeepSeek 搜索服务异常（HTTP {response.status_code}）{f'：{detail[:120]}' if detail else ''}")
    return _map_tools_search_results(response.json()), "", "deepseek-search-api"


async def _search_single_anthropic(
    client: httpx.AsyncClient,
    cfg: dict[str, Any],
    query: str,
) -> tuple[list[dict[str, str]], str, str]:
    endpoint = f"{cfg['anthropic_url']}/messages"
    body = {
        "model": cfg["model"],
        "max_tokens": 4096,
        "messages": [
            {
                "role": "user",
                "content": [{"type": "text", "text": f"Perform a web search for the query: {query}"}],
            }
        ],
        "tools": [
            {
                "type": "web_search_20250305",
                "name": "web_search",
                "max_uses": 5,
            }
        ],
        "thinking": {"type": "disabled"},
    }
    response = await client.post(
        endpoint,
        headers={
            "x-api-key": cfg["api_key"],
            "Authorization": f"Bearer {cfg['api_key']}",
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": USER_AGENT,
        },
        json=body,
    )
    if response.status_code == 401:
        fail(502, "DeepSeek 搜索密钥无效或已过期，请检查 QO_DEEPSEEK_SEARCH_API_KEY")
    if response.status_code != 200:
        LOG.warning("DeepSeek anthropic fallback search failed HTTP %s", response.status_code)
        fail(502, f"DeepSeek 联网搜索失败（HTTP {response.status_code}）")
    sources, answer = _map_anthropic_search_results(response.json())
    return sources, answer, "deepseek-official"


@router.get("/status", dependencies=[Depends(user_for)])
def get_search_status():
    return search_status()


@router.post("/query")
async def query_web_search(body: SearchQueryRequest, user: dict = Depends(user_for)):
    cfg = _config()
    if not cfg["api_key"]:
        fail(503, "服务端尚未配置 DeepSeek 搜索密钥（QO_DEEPSEEK_SEARCH_API_KEY）")
    cleaned_queries: list[str] = []
    seen_q: set[str] = set()
    for raw in body.queries:
        q = " ".join(str(raw or "").split()).strip()
        if q and q not in seen_q:
            seen_q.add(q)
            cleaned_queries.append(q[:200])
    if not cleaned_queries:
        fail(422, "请提供至少一个非空搜索关键词")

    ori_question = _bounded_question(body.ori_question)
    try:
        async with httpx.AsyncClient(timeout=35.0, follow_redirects=False) as client:
            outcomes = await asyncio.gather(
                *[
                    _search_single_tools(client, cfg, q, body.count, ori_question, body.need_content)
                    for q in cleaned_queries
                ]
            )
    except httpx.HTTPError as exc:
        LOG.warning("DeepSeek search network error: %s", exc)
        fail(502, "无法连接 DeepSeek 搜索服务，请稍后重试")

    seen_urls: set[str] = set()
    merged: list[dict[str, str]] = []
    answers: list[str] = []
    provider = "deepseek-search-api"
    max_rank = max((len(item[0]) for item in outcomes), default=0)
    truncated = False
    for rank in range(max_rank):
        for sources, _answer, prov in outcomes:
            provider = prov
            if rank < len(sources):
                src = sources[rank]
                if src["url"] not in seen_urls:
                    seen_urls.add(src["url"])
                    if len(merged) >= body.count:
                        truncated = True
                    else:
                        merged.append(src)
    for idx, (_sources, answer, _prov) in enumerate(outcomes):
        if answer:
            answers.append(f"### {cleaned_queries[idx]}\n{answer}" if len(cleaned_queries) > 1 else answer)

    try:
        with database(write=True) as db:
            audit(db, "web.search", user["id"], None, queries=cleaned_queries, results=len(merged), provider=provider)
    except Exception:
        pass

    return {
        "ok": True,
        "provider": provider,
        "queries": cleaned_queries,
        "sources": merged,
        "answer": "\n\n".join(answers),
        "truncated": truncated,
    }


@router.post("/fetch")
async def fetch_web_page(body: FetchUrlRequest, user: dict = Depends(user_for)):
    cfg = _config()
    if not cfg["api_key"]:
        fail(503, "服务端尚未配置 DeepSeek 搜索密钥（QO_DEEPSEEK_SEARCH_API_KEY）")
    target_url = body.url.strip()
    if not _is_public_http_url(target_url):
        fail(422, "仅允许抓取公开的 HTTP/HTTPS 网页地址")

    endpoint = f"{cfg['base_url']}/fetch"
    try:
        async with httpx.AsyncClient(timeout=35.0, follow_redirects=False) as client:
            response = await client.post(
                endpoint,
                headers={
                    "Authorization": f"Bearer {cfg['api_key']}",
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                    "User-Agent": USER_AGENT,
                },
                json={"url": target_url},
            )
    except httpx.HTTPError as exc:
        LOG.warning("DeepSeek fetch network error: %s", exc)
        fail(502, "无法连接 DeepSeek 网页抓取服务，请稍后重试")

    if response.status_code == 401:
        fail(502, "DeepSeek 搜索密钥无效或已过期，请检查 QO_DEEPSEEK_SEARCH_API_KEY")
    if response.status_code != 200:
        fail(502, f"网页抓取失败（HTTP {response.status_code}）")

    payload = response.json()
    content = payload.get("content") if isinstance(payload, dict) else None
    if not isinstance(content, str) or not content.strip():
        fail(502, "该网页未返回可提取的正文内容")

    truncated = len(content) > MAX_FETCH_CHARS
    trimmed = content[:MAX_FETCH_CHARS]
    try:
        with database(write=True) as db:
            audit(db, "web.fetch", user["id"], None, url=target_url[:300], chars=len(trimmed))
    except Exception:
        pass

    return {
        "ok": True,
        "url": target_url,
        "content": trimmed,
        "truncated": truncated,
    }
