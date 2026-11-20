"use strict";

const { createServer } = require("node:http");
const { randomBytes, timingSafeEqual } = require("node:crypto");

const MAX_BODY_BYTES = 32 * 1024;
const SEARCH_TIMEOUT_MS = 40_000;

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

async function createSearchBroker({ search, fetchPage, onSearch = () => {}, onFetch = () => {} }) {
  if (typeof search !== "function" || typeof fetchPage !== "function") {
    throw new TypeError("搜索 broker 参数无效");
  }
  const key = randomBytes(32).toString("base64url");
  const secret = Buffer.from(key, "utf8");
  let searchCount = 0;
  let fetchCount = 0;
  const queriesHistory = [];

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
        const queries = [...new Set(
          rawQueries
            .map((item) => (typeof item === "string" ? item.replace(/\s+/g, " ").trim() : ""))
            .filter(Boolean)
            .slice(0, 4)
        )];
        if (!queries.length) {
          json(response, 400, { ok: false, message: "请提供至少 1 个非空搜索关键词" });
          return;
        }
        const count = Number.isInteger(body.value.count) && body.value.count >= 1 && body.value.count <= 20
          ? body.value.count
          : 8;
        const oriQuestion = typeof body.value.ori_question === "string" ? body.value.ori_question.trim().slice(0, 200) : undefined;
        const needContent = body.value.need_content === true;
        onSearch(queries);
        const result = await search({ queries, count, ori_question: oriQuestion, need_content: needContent }, { signal: controller.signal });
        searchCount += 1;
        queriesHistory.push(...queries);
        json(response, 200, { ok: true, ...result });
      } else {
        const targetUrl = typeof body.value.url === "string" ? body.value.url.trim() : "";
        if (!targetUrl || !/^https?:\/\//i.test(targetUrl)) {
          json(response, 400, { ok: false, message: "请提供有效的 HTTP/HTTPS 网址" });
          return;
        }
        onFetch(targetUrl);
        const result = await fetchPage({ url: targetUrl }, { signal: controller.signal });
        fetchCount += 1;
        json(response, 200, { ok: true, ...result });
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
    queriesHistory,
    async close() {
      if (!server.listening) return;
      await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    }
  };
}

module.exports = { MAX_BODY_BYTES, SEARCH_TIMEOUT_MS, createSearchBroker };
