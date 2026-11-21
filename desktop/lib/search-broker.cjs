"use strict";

const { createServer } = require("node:http");
const { randomBytes, timingSafeEqual } = require("node:crypto");

const MAX_BODY_BYTES = 32 * 1024;
const SEARCH_TIMEOUT_MS = 40_000;
const YEAR_ONLY_RE = /\b(202[0-9])\b/g;
const HISTORICAL_MONTH_RE = /202\d\s*年\s*\d{1,2}\s*月/;
const HIST_BG_STOPWORDS_RE = /\b(?:202\d(?:年(?:\d{1,2}月)?)?|成立时间|创始人|公司简介|发展历程|融资历程|背景)\b/g;

function normalizeServerTime(serverTime, now = Date.now) {
  if (serverTime && typeof serverTime === "object" && Number.isInteger(serverTime.unix_timestamp) && serverTime.year && serverTime.date) {
    return {
      unix_timestamp: serverTime.unix_timestamp,
      year: Number(serverTime.year),
      month: Number(serverTime.month) || 1,
      date: String(serverTime.date),
      display: String(serverTime.display || `${serverTime.date} (Unix时间戳: ${serverTime.unix_timestamp})`)
    };
  }
  const ms = typeof now === "function" ? now() : Date.now();
  const shanghaiDate = new Date(ms + 8 * 3600 * 1000);
  const year = shanghaiDate.getUTCFullYear();
  const month = shanghaiDate.getUTCMonth() + 1;
  const day = shanghaiDate.getUTCDate();
  const hh = String(shanghaiDate.getUTCHours()).padStart(2, "0");
  const mm = String(shanghaiDate.getUTCMinutes()).padStart(2, "0");
  const ss = String(shanghaiDate.getUTCSeconds()).padStart(2, "0");
  const unix_timestamp = Math.floor(ms / 1000);
  const date = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  return {
    unix_timestamp,
    year,
    month,
    date,
    display: `${year}年${String(month).padStart(2, "0")}月${String(day).padStart(2, "0")}日 ${hh}:${mm}:${ss} (UTC+8, Unix时间戳: ${unix_timestamp})`
  };
}

function anchorQueriesToServerTime(queries, srvTime) {
  const curYear = Number(srvTime?.year) || new Date().getFullYear();
  const curYearStr = String(curYear);
  const staleYears = new Set([String(curYear - 1), String(curYear - 2)]);
  const anchored = [];
  const seen = new Set();

  for (const q of queries) {
    let updated = q;
    const foundYears = new Set([...updated.matchAll(YEAR_ONLY_RE)].map((m) => m[1]));
    const hasStale = [...foundYears].some((y) => staleYears.has(y));
    const hasCurrent = foundYears.has(curYearStr);
    if (hasStale && !hasCurrent) {
      const hasExplicitHistMonth = HISTORICAL_MONTH_RE.test(updated);
      const hasMultiYearCompare = foundYears.size >= 2 || updated.includes("年报");
      if (!hasExplicitHistMonth && !hasMultiYearCompare) {
        for (const sy of staleYears) {
          updated = updated.replace(new RegExp(`\\b${sy}(?=\\s*年|\\b)`, "g"), curYearStr);
        }
      } else if (updated.includes("最新") || updated.includes("近况") || updated.includes("动态")) {
        updated = `${updated} ${curYearStr}`;
      }
    }
    updated = updated.replace(/\s+/g, " ").trim().slice(0, 200);
    if (updated && !seen.has(updated)) {
      seen.add(updated);
      anchored.push(updated);
    }
  }

  const rawHadCurrentYear = queries.some((q) => q.includes(curYearStr));
  if (anchored.length > 0 && (!rawHadCurrentYear || !anchored.some((q) => q.includes(curYearStr)))) {
    const first = anchored[0];
    const cleanedTopic = first.replace(HIST_BG_STOPWORDS_RE, " ");
    const tokens = cleanedTopic.split(/\s+/).filter((tok) => tok && !/^abab[\d.]*$/i.test(tok));
    const topic = tokens.slice(0, 2).join(" ").trim() || first.split(/\s+/)[0];
    const freshQuery = `${topic} ${curYearStr}年 最新模型 产品发布 动态`.slice(0, 200);
    if (!seen.has(freshQuery)) {
      anchored.unshift(freshQuery);
    }
  }

  return anchored.slice(0, 5);
}

function json(response, status, value) {
  const body = Buffer.from(JSON.stringify(value));
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": body.length,
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff"
  });
  response.end(body);
}

function equalSecret(received, expected) {
  const actual = Buffer.from(received || "", "utf8");
  const padded = Buffer.alloc(expected.length);
  actual.copy(padded, 0, 0, expected.length);
  return timingSafeEqual(expected, padded) && actual.length === expected.length;
}

async function readBody(request) {
  const contentLength = Number(request.headers["content-length"]);
  if (Number.isFinite(contentLength) && contentLength > MAX_BODY_BYTES) return { tooLarge: true };
  const chunks = [];
  let length = 0;
  for await (const chunk of request) {
    length += chunk.length;
    if (length > MAX_BODY_BYTES) return { tooLarge: true };
    chunks.push(chunk);
  }
  try {
    return { value: JSON.parse(Buffer.concat(chunks).toString("utf8")) };
  } catch {
    return { invalid: true };
  }
}

async function createSearchBroker({
  search,
  fetchPage,
  serverTime,
  now = Date.now,
  initialState = null,
  onSearch = () => {},
  onFetch = () => {},
  onStateChange = () => {}
}) {
  if (typeof search !== "function" || typeof fetchPage !== "function") {
    throw new TypeError("搜索 broker 参数无效");
  }
  const key = randomBytes(32).toString("base64url");
  const secret = Buffer.from(key, "utf8");
  let activeServerTime = normalizeServerTime(serverTime || initialState?.serverTime, now);
  let searchCount = Number.isInteger(initialState?.searchCount) && initialState.searchCount > 0 ? initialState.searchCount : 0;
  let fetchCount = Number.isInteger(initialState?.fetchCount) && initialState.fetchCount > 0 ? initialState.fetchCount : 0;
  const queriesHistory = Array.isArray(initialState?.queriesHistory)
    ? initialState.queriesHistory.filter((q) => typeof q === "string" && q.trim()).slice(0, 40)
    : [];
  const evidenceByUrl = new Map();
  if (Array.isArray(initialState?.evidenceByUrl)) {
    for (const item of initialState.evidenceByUrl.slice(0, 30)) {
      if (item && typeof item.url === "string" && item.url) {
        evidenceByUrl.set(item.url, {
          url: item.url,
          title: typeof item.title === "string" && item.title.trim() ? item.title.trim() : item.url,
          published_at: typeof item.published_at === "string" ? item.published_at.trim() : "",
          excerpt: typeof item.excerpt === "string" ? item.excerpt.slice(0, 320) : ""
        });
      }
    }
  }
  const fetchedPages = new Map();
  if (Array.isArray(initialState?.fetchedPages)) {
    for (const entry of initialState.fetchedPages.slice(0, 10)) {
      if (Array.isArray(entry) && typeof entry[0] === "string" && typeof entry[1] === "string" && entry[0]) {
        fetchedPages.set(entry[0], entry[1].slice(0, 520));
      }
    }
  }

  function exportState() {
    return {
      searchCount,
      fetchCount,
      queriesHistory: [...queriesHistory].slice(-40),
      evidenceByUrl: [...evidenceByUrl.values()].slice(0, 30),
      fetchedPages: [...fetchedPages.entries()].slice(0, 10),
      serverTime: activeServerTime
    };
  }

  function recordSearchSources(sources = []) {
    if (!Array.isArray(sources)) return;
    for (const src of sources.slice(0, 8)) {
      if (!src || typeof src.url !== "string" || !src.url) continue;
      const excerpt = (typeof src.summary === "string" && src.summary.trim())
        ? src.summary.trim().replace(/\s+/g, " ").slice(0, 320)
        : (typeof src.snippet === "string" ? src.snippet.trim().replace(/\s+/g, " ").slice(0, 320) : "");
      evidenceByUrl.set(src.url, {
        url: src.url,
        title: typeof src.title === "string" && src.title.trim() ? src.title.trim() : src.url,
        published_at: typeof src.published_at === "string" ? src.published_at.trim() : "",
        excerpt
      });
    }
  }

  function getEvidenceDigest(maxChars = 6000) {
    if (!evidenceByUrl.size && !fetchedPages.size) return "";
    const lines = [
      `[服务器权威时间基准] ${activeServerTime.display} (当前年份: ${activeServerTime.year}年, 基准日期: ${activeServerTime.date}, Unix时间戳: ${activeServerTime.unix_timestamp})`,
      `[本轮已执行的检索词] ${[...new Set(queriesHistory)].slice(0, 16).join(" | ")}`
    ];
    if (evidenceByUrl.size) {
      lines.push("【已检索核实的最新权威来源摘要】");
      for (const item of [...evidenceByUrl.values()].slice(0, 14)) {
        const dateTag = item.published_at ? `[${item.published_at}] ` : "";
        lines.push(`- ${dateTag}${item.title} (${item.url})${item.excerpt ? `：${item.excerpt}` : ""}`);
      }
    }
    if (fetchedPages.size) {
      lines.push("【已精读的网页正文摘要】");
      for (const [url, snippet] of [...fetchedPages.entries()].slice(0, 4)) {
        lines.push(`- ${url}：${snippet}`);
      }
    }
    return lines.join("\n").slice(0, maxChars);
  }

  const server = createServer(async (request, response) => {
    if (request.method !== "POST" || (request.url !== "/search" && request.url !== "/fetch")) {
      json(response, 404, { ok: false, message: "未找到请求路径" });
      request.resume();
      return;
    }
    const authorization = request.headers.authorization || "";
    const bearer = authorization.startsWith("Bearer ") ? authorization.slice(7) : "";
    if (!equalSecret(bearer, secret)) {
      json(response, 401, { ok: false, message: "搜索工具认证失败" });
      request.resume();
      return;
    }
    const body = await readBody(request);
    if (body.tooLarge) {
      json(response, 413, { ok: false, message: "搜索请求超过大小限制" });
      return;
    }
    if (body.invalid || !body.value || typeof body.value !== "object" || Array.isArray(body.value)) {
      json(response, 400, { ok: false, message: "搜索请求格式无效" });
      return;
    }

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(new Error("搜索超时")), SEARCH_TIMEOUT_MS);
    timer.unref?.();
    request.once("aborted", () => controller.abort(new Error("搜索请求已取消")));
    response.once("close", () => {
      if (!response.writableEnded) controller.abort(new Error("搜索请求已取消"));
    });

    try {
      if (request.url === "/search") {
        const rawQueries = Array.isArray(body.value.queries)
          ? body.value.queries
          : typeof body.value.query === "string"
            ? [body.value.query]
            : [];
        const cleanedQueries = [...new Set(
          rawQueries
            .map((item) => (typeof item === "string" ? item.replace(/\s+/g, " ").trim() : ""))
            .filter(Boolean)
            .slice(0, 4)
        )];
        if (!cleanedQueries.length) {
          json(response, 400, { ok: false, message: "请提供至少 1 个非空搜索关键词" });
          return;
        }
        const queries = anchorQueriesToServerTime(cleanedQueries, activeServerTime);
        const count = Number.isInteger(body.value.count) && body.value.count >= 1 && body.value.count <= 20
          ? body.value.count
          : 8;
        const oriQuestion = typeof body.value.ori_question === "string" ? body.value.ori_question.trim().slice(0, 200) : undefined;
        const needContent = body.value.need_content === true;
        onSearch(queries, activeServerTime);
        const result = await search({ queries, count, ori_question: oriQuestion, need_content: needContent }, { signal: controller.signal });
        if (result?.server_time && typeof result.server_time === "object") {
          activeServerTime = normalizeServerTime(result.server_time, now);
        }
        const effectiveQueries = Array.isArray(result?.queries) && result.queries.length ? result.queries : queries;
        searchCount += 1;
        queriesHistory.push(...effectiveQueries);
        recordSearchSources(result?.sources);
        try { onStateChange(exportState()); } catch {}
        json(response, 200, { ok: true, server_time: activeServerTime, ...result, queries: effectiveQueries });
      } else {
        const targetUrl = typeof body.value.url === "string" ? body.value.url.trim() : "";
        if (!targetUrl || !/^https?:\/\//i.test(targetUrl)) {
          json(response, 400, { ok: false, message: "请提供有效的 HTTP/HTTPS 网址" });
          return;
        }
        onFetch(targetUrl);
        const result = await fetchPage({ url: targetUrl }, { signal: controller.signal });
        if (result?.server_time && typeof result.server_time === "object") {
          activeServerTime = normalizeServerTime(result.server_time, now);
        }
        fetchCount += 1;
        if (typeof result?.content === "string" && result.content.trim()) {
          fetchedPages.set(targetUrl, result.content.trim().replace(/\s+/g, " ").slice(0, 520));
        }
        try { onStateChange(exportState()); } catch {}
        json(response, 200, { ok: true, server_time: activeServerTime, ...result });
      }
    } catch (error) {
      const raw = typeof error?.message === "string" ? error.message : "搜索服务暂不可用";
      const message = raw.replaceAll(key, "[已隐藏]").slice(0, 300);
      json(response, 502, { ok: false, message: message || "搜索服务暂不可用" });
    } finally {
      clearTimeout(timer);
    }
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  return {
    url: `http://127.0.0.1:${address.port}`,
    key,
    get searchCount() { return searchCount; },
    get fetchCount() { return fetchCount; },
    get serverTime() { return activeServerTime; },
    queriesHistory,
    getEvidenceDigest,
    exportState,
    async close() {
      if (!server.listening) return;
      await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    }
  };
}

module.exports = { MAX_BODY_BYTES, SEARCH_TIMEOUT_MS, anchorQueriesToServerTime, createSearchBroker, normalizeServerTime };
