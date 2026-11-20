"""DeepSeek 官方 Search & Fetch API 服务端代理（集中保管密钥，供桌面 Harness 回环调用）。"""

from __future__ import annotations

import asyncio
import ipaddress
import logging
import os
import re
from datetime import datetime
from typing import Any
from urllib.parse import urlsplit

import httpx
from fastapi import APIRouter, Depends
from pydantic import BaseModel, ConfigDict, Field

from .identity import fail, user_for
from .store import ZONE, audit, database

router = APIRouter(prefix="/api/v1/search")
LOG = logging.getLogger("qiongyu.search")

DEFAULT_TOOLS_BASE_URL = "https://api.deepseek.com/v1/tools"
DEFAULT_ANTHROPIC_BASE_URL = "https://api.deepseek.com/anthropic/v1"
DEFAULT_ANTHROPIC_MODEL = "deepseek-v4-flash"
USER_AGENT = "deepseek-harness/0.0.1"
MAX_FETCH_CHARS = 24_000
WEEKDAYS_CN = ("星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日")
_MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
_ISO_DATE_RE = re.compile(r"\b(202\d)[-/](\d{1,2})[-/](\d{1,2})\b")
_CN_DATE_RE = re.compile(r"(202\d)\s*年\s*(\d{1,2})\s*月(?:\s*(\d{1,2})\s*日)?")
_EN_DATE_RE = re.compile(
    r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z.]*\s+(\d{1,2}),?\s+(202\d)\b",
    re.IGNORECASE,
)
_EN_MONTH_YEAR_RE = re.compile(
    r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z.]*\s+(202\d)\b",
    re.IGNORECASE,
)
_YEAR_ONLY_RE = re.compile(r"\b(202[0-9])\b")
_HISTORICAL_MONTH_RE = re.compile(r"202\d\s*年\s*\d{1,2}\s*月")
_HIST_BG_STOPWORDS = re.compile(r"\b(?:202\d(?:年(?:\d{1,2}月)?)?|成立时间|创始人|公司简介|发展历程|融资历程|背景)\b")


def server_time_snapshot(now: datetime | None = None) -> dict[str, Any]:
    """返回当前服务器的权威时间与 Unix 时间戳快照，供搜索词校准与提示词基准注入。"""
    current = (now or datetime.now(ZONE)).astimezone(ZONE)
    unix_ts = int(current.timestamp())
    weekday = WEEKDAYS_CN[current.weekday()]
    display = (
        f"{current.year}年{current.month:02d}月{current.day:02d}日 {weekday} "
        f"{current.strftime('%H:%M:%S')} (UTC+8, Unix时间戳: {unix_ts})"
    )
    return {
        "unix_timestamp": unix_ts,
        "unix_ms": int(current.timestamp() * 1000),
        "iso": current.isoformat(timespec="seconds"),
        "date": current.strftime("%Y-%m-%d"),
        "year": current.year,
        "month": current.month,
        "day": current.day,
        "weekday": weekday,
        "timezone": str(ZONE),
        "display": display,
    }


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
        "server_time": server_time_snapshot(),
    }


class SearchQueryRequest(BaseModel):
    model_config = ConfigDict(extra="ignore")
    queries: list[str] = Field(min_length=1, max_length=5)
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


def _safe_date_ts(year: int, month: int, day: int = 15) -> tuple[str, int] | None:
    if not (2000 <= year <= 2035 and 1 <= month <= 12):
        return None
    safe_day = max(1, min(28 if month == 2 else 30 if month in {4, 6, 9, 11} else 31, day))
    try:
        dt = datetime(year, month, safe_day, 12, 0, 0, tzinfo=ZONE)
        return f"{year:04d}-{month:02d}-{safe_day:02d}", int(dt.timestamp())
    except ValueError:
        return None


def _extract_source_timestamp(item: dict[str, Any], srv_time: dict[str, Any]) -> tuple[str, int]:
    """根据 DeepSeek 搜索结果的 date 字段及标题/摘要中的日期提取最新合法时间戳。"""
    max_allowed_ts = int(srv_time["unix_timestamp"]) + 30 * 86400
    raw_date = str(item.get("date") or "").strip()
    candidates: list[tuple[str, int]] = []

    if raw_date:
        try:
            iso_norm = raw_date.replace("Z", "+00:00")
            parsed_dt = datetime.fromisoformat(iso_norm)
            if parsed_dt.tzinfo is None:
                parsed_dt = parsed_dt.replace(tzinfo=ZONE)
            ts = int(parsed_dt.timestamp())
            if ts <= max_allowed_ts:
                return parsed_dt.astimezone(ZONE).strftime("%Y-%m-%d"), ts
        except ValueError:
            pass
        for m in _ISO_DATE_RE.finditer(raw_date):
            hit = _safe_date_ts(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if hit and hit[1] <= max_allowed_ts:
                return hit

    combined = "\n".join(
        str(item.get(k) or "") for k in ("title", "snippet", "summary")
    )[:2400]
    for m in _ISO_DATE_RE.finditer(combined):
        hit = _safe_date_ts(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if hit and hit[1] <= max_allowed_ts:
            candidates.append(hit)
    for m in _CN_DATE_RE.finditer(combined):
        day = int(m.group(3)) if m.group(3) else 15
        hit = _safe_date_ts(int(m.group(1)), int(m.group(2)), day)
        if hit and hit[1] <= max_allowed_ts:
            candidates.append(hit)
    for m in _EN_DATE_RE.finditer(combined):
        month = _MONTH_MAP.get(m.group(1)[:3].lower())
        if month:
            hit = _safe_date_ts(int(m.group(3)), month, int(m.group(2)))
            if hit and hit[1] <= max_allowed_ts:
                candidates.append(hit)
    for m in _EN_MONTH_YEAR_RE.finditer(combined):
        month = _MONTH_MAP.get(m.group(1)[:3].lower())
        if month:
            hit = _safe_date_ts(int(m.group(2)), month, 15)
            if hit and hit[1] <= max_allowed_ts:
                candidates.append(hit)

    if candidates:
        return max(candidates, key=lambda pair: pair[1])
    if raw_date:
        return raw_date[:32], 0
    return "", 0


def _anchor_queries_to_server_time(queries: list[str], srv_time: dict[str, Any]) -> list[str]:
    """根据当前服务器时间戳校准查询词：自动修正模型因训练语料滞后写出的旧年份，并确保批次覆盖当前年份最新进展。"""
    cur_year = int(srv_time["year"])
    cur_year_str = str(cur_year)
    stale_years = {str(cur_year - 1), str(cur_year - 2)}
    anchored: list[str] = []
    seen: set[str] = set()

    for q in queries:
        updated = q
        found_years = set(_YEAR_ONLY_RE.findall(updated))
        has_stale = bool(found_years & stale_years)
        has_current = cur_year_str in found_years
        if has_stale and not has_current:
            has_explicit_hist_month = bool(_HISTORICAL_MONTH_RE.search(updated))
            has_multi_year_compare = len(found_years) >= 2 or "年报" in updated
            if not has_explicit_hist_month and not has_multi_year_compare:
                for sy in stale_years:
                    updated = re.sub(rf"\b{sy}(?=\s*年|\b)", cur_year_str, updated)
            elif "最新" in updated or "近况" in updated or "动态" in updated:
                updated = f"{updated} {cur_year_str}"
        updated = " ".join(updated.split()).strip()[:200]
        if updated and updated not in seen:
            seen.add(updated)
            anchored.append(updated)

    raw_had_current_year = any(cur_year_str in q for q in queries)
    if anchored and (not raw_had_current_year or not any(cur_year_str in q for q in anchored)):
        first = anchored[0]
        cleaned_topic = _HIST_BG_STOPWORDS.sub(" ", first)
        tokens = [tok for tok in cleaned_topic.split() if tok and not re.fullmatch(r"abab[\d.]*", tok, re.I)]
        topic = " ".join(tokens[:2]).strip() or first.split()[0]
        fresh_query = f"{topic} {cur_year_str}年 最新模型 产品发布 动态"[:200]
        if fresh_query not in seen:
            anchored = [fresh_query, *anchored][:5]

    return anchored or queries


def _map_tools_search_results(payload: Any, srv_time: dict[str, Any]) -> list[dict[str, Any]]:
    if not isinstance(payload, dict) or not isinstance(payload.get("results"), list):
        raise ValueError("DeepSeek Search API 返回结构无效")
    sources: list[dict[str, Any]] = []
    for item in payload["results"]:
        if not isinstance(item, dict):
            continue
        url = str(item.get("url") or "").strip()
        if not url:
            continue
        entry: dict[str, Any] = {"url": url}
        title = str(item.get("title") or "").strip()
        snippet = str(item.get("snippet") or "").strip()
        summary = str(item.get("summary") or "").strip()
        date_label, date_ts = _extract_source_timestamp(item, srv_time)
        if title:
            entry["title"] = title
        if snippet:
            entry["snippet"] = snippet
        if summary:
            entry["summary"] = summary
        if date_label:
            entry["published_at"] = date_label
        if date_ts > 0:
            entry["published_ts"] = date_ts
        sources.append(entry)
    return sources


def _map_anthropic_search_results(payload: Any) -> tuple[list[dict[str, Any]], str]:
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
    sources: list[dict[str, Any]] = []
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
            entry: dict[str, Any] = {"url": url}
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
    srv_time: dict[str, Any],
) -> tuple[list[dict[str, Any]], str, str]:
    endpoint = f"{cfg['base_url']}/search"
    time_tagged_question = _bounded_question(
        f"[当前服务器时间:{srv_time['date']} timestamp:{srv_time['unix_timestamp']}] {ori_question or query}"
    )
    body: dict[str, Any] = {
        "query": query[:128],
        "count": count,
        "need_content": bool(need_content),
    }
    if time_tagged_question:
        body["ori_question"] = time_tagged_question
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
        return await _search_single_anthropic(client, cfg, query, srv_time)
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
    return _map_tools_search_results(response.json(), srv_time), "", "deepseek-search-api"


async def _search_single_anthropic(
    client: httpx.AsyncClient,
    cfg: dict[str, Any],
    query: str,
    srv_time: dict[str, Any],
) -> tuple[list[dict[str, Any]], str, str]:
    endpoint = f"{cfg['anthropic_url']}/messages"
    body = {
        "model": cfg["model"],
        "max_tokens": 4096,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            f"Current server time: {srv_time['iso']} (Unix timestamp: {srv_time['unix_timestamp']}). "
                            f"Perform a web search prioritizing the latest {srv_time['year']} information for: {query}"
                        ),
                    }
                ],
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


def _freshness_score(src: dict[str, Any], rank: int, query_idx: int, srv_time: dict[str, Any]) -> float:
    now_ts = int(srv_time["unix_timestamp"])
    cur_year_str = str(srv_time["year"])
    prev_year_str = str(int(srv_time["year"]) - 1)
    pub_ts = int(src.get("published_ts") or 0)
    pub_label = str(src.get("published_at") or "")
    bonus = 0.0
    if pub_ts > 0:
        age_days = max(0.0, (now_ts - pub_ts) / 86400.0)
        if age_days <= 90:
            bonus = 4.5
        elif age_days <= 240:
            bonus = 3.2
        elif pub_label.startswith(cur_year_str):
            bonus = 2.5
        elif pub_label.startswith(prev_year_str):
            bonus = 0.5
        else:
            bonus = -1.2
    elif pub_label.startswith(cur_year_str):
        bonus = 2.5
    elif pub_label.startswith(prev_year_str):
        bonus = 0.5
    return bonus - (rank * 0.45) - (query_idx * 0.15)


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

    srv_time = server_time_snapshot()
    effective_queries = _anchor_queries_to_server_time(cleaned_queries, srv_time)
    ori_question = _bounded_question(body.ori_question)
    per_query_count = max(body.count, 8)
    try:
        async with httpx.AsyncClient(timeout=35.0, follow_redirects=False) as client:
            outcomes = await asyncio.gather(
                *[
                    _search_single_tools(
                        client, cfg, q, per_query_count, ori_question, body.need_content, srv_time
                    )
                    for q in effective_queries
                ]
            )
    except httpx.HTTPError as exc:
        LOG.warning("DeepSeek search network error: %s", exc)
        fail(502, "无法连接 DeepSeek 搜索服务，请稍后重试")

    by_url: dict[str, tuple[float, dict[str, Any]]] = {}
    answers: list[str] = []
    provider = "deepseek-search-api"
    for q_idx, (sources, answer, prov) in enumerate(outcomes):
        provider = prov
        if answer:
            answers.append(f"### {effective_queries[q_idx]}\n{answer}" if len(effective_queries) > 1 else answer)
        for rank, src in enumerate(sources):
            url = src["url"]
            score = _freshness_score(src, rank, q_idx, srv_time)
            existing = by_url.get(url)
            if existing is None or score > existing[0]:
                by_url[url] = (score, src)

    ranked_sources = [item[1] for item in sorted(by_url.values(), key=lambda pair: pair[0], reverse=True)]
    truncated = len(ranked_sources) > body.count
    merged = ranked_sources[: body.count]

    try:
        with database(write=True) as db:
            audit(
                db,
                "web.search",
                user["id"],
                None,
                queries=effective_queries,
                raw_queries=cleaned_queries if effective_queries != cleaned_queries else None,
                server_timestamp=srv_time["unix_timestamp"],
                results=len(merged),
                provider=provider,
            )
    except Exception:
        pass

    return {
        "ok": True,
        "provider": provider,
        "server_time": srv_time,
        "queries": effective_queries,
        "raw_queries": cleaned_queries,
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
        "server_time": server_time_snapshot(),
        "url": target_url,
        "content": trimmed,
        "truncated": truncated,
    }
