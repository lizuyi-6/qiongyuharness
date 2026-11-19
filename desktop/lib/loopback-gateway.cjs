const http = require("node:http");
const { randomBytes, timingSafeEqual } = require("node:crypto");

const REQUEST_LIMIT = 2 * 1024 * 1024;
const RESPONSE_LIMIT = 3 * 1024 * 1024;
const TOKEN_LIMIT = 3500;
const MAX_CALLS = 4;
const FORWARDED_FIELDS = [
  "model", "messages", "stream", "stream_options", "temperature", "top_p",
  "max_tokens", "reasoning_effort", "thinking"
];

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function bearerMatches(header, key) {
  const expected = Buffer.from(`Bearer ${key}`);
  const actual = Buffer.from(typeof header === "string" ? header : "");
  return actual.length === expected.length && timingSafeEqual(actual, expected);
}

function allowedFunction(value, allowed) {
  return isObject(value) && typeof value.name === "string" && allowed.has(value.name);
}

function sanitizeRequest(value, model, allowedTools) {
  if (!isObject(value) || value.model !== model || !Array.isArray(value.messages)) {
    throw new Error("invalid request");
  }

  const tools = value.tools ?? [];
  if (!Array.isArray(tools) || tools.some((tool) =>
    !isObject(tool) || tool.type !== "function" || !allowedFunction(tool.function, allowedTools))) {
    throw new Error("invalid tools");
  }

  const roles = new Set(["system", "developer", "user", "assistant"]);
  if (allowedTools.size) roles.add("tool");
  for (const message of value.messages) {
    if (!isObject(message) || !roles.has(message.role)) throw new Error("invalid message role");
    if (message.tool_calls !== undefined && (!Array.isArray(message.tool_calls) || message.tool_calls.some((call) =>
      !isObject(call) || !isObject(call.function) || !allowedFunction(call.function, allowedTools)))) {
      throw new Error("invalid tool call");
    }
  }

  const params = {};
  for (const field of FORWARDED_FIELDS) {
    if (Object.hasOwn(value, field)) params[field] = value[field];
  }
  if (tools.length) params.tools = tools;

  let maxTokens = value.max_tokens;
  if (maxTokens === undefined || maxTokens === null || maxTokens === 0) maxTokens = TOKEN_LIMIT;
  if (typeof maxTokens !== "number" || !Number.isFinite(maxTokens) || !Number.isInteger(maxTokens)) {
    throw new Error("invalid max_tokens");
  }
  params.max_tokens = Math.min(Math.max(maxTokens, 1), TOKEN_LIMIT);
  return params;
}

async function readBody(request) {
  const length = request.headers["content-length"];
  if (length !== undefined && (!/^\d+$/.test(length) || Number(length) > REQUEST_LIMIT)) {
    const error = new Error("request body too large");
    error.statusCode = 413;
    throw error;
  }
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > REQUEST_LIMIT) {
      const error = new Error("request body too large");
      error.statusCode = 413;
      throw error;
    }
    chunks.push(chunk);
  }
  if (size === 0) throw new Error("empty request");
  return Buffer.concat(chunks, size);
}

function usageFromPayload(payload) {
  const usage = payload?.usage;
  if (!isObject(usage)) return {};
  return Object.fromEntries(Object.entries(usage).filter(([, value]) =>
    Number.isSafeInteger(value) && value >= 0));
}

function usageFromResponse(contentType, body) {
  if (contentType.includes("text/event-stream")) {
    let latest = {};
    for (const line of body.toString("utf8").split(/\r?\n/)) {
      if (!line.startsWith("data:")) continue;
      const data = line.slice(5).trim();
      if (!data || data === "[DONE]") continue;
      try {
        latest = usageFromPayload(JSON.parse(data)) || latest;
      } catch {}
    }
    return latest;
  }
  try {
    return usageFromPayload(JSON.parse(body.toString("utf8")));
  } catch {
    return {};
  }
}

function addUsage(target, next) {
  for (const [key, value] of Object.entries(next)) {
    target[key] = (target[key] || 0) + value;
  }
}

function replyError(response, status, message) {
  if (response.headersSent || response.destroyed) return;
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
  response.end(JSON.stringify({ error: { message } }));
}

async function pipeResponse(upstream, response, state) {
  const contentType = upstream.headers.get("content-type") || "application/json";
  if (!upstream.ok || !upstream.body) {
    replyError(response, 502, "模型服务请求失败");
    return;
  }

  response.writeHead(upstream.status, {
    "Content-Type": contentType,
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff"
  });
  const reader = upstream.body.getReader();
  const chunks = [];
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > RESPONSE_LIMIT) {
        await reader.cancel();
        response.destroy();
        return;
      }
      const chunk = Buffer.from(value);
      chunks.push(chunk);
      if (!response.write(chunk)) await new Promise((resolve) => response.once("drain", resolve));
    }
    response.end();
    addUsage(state.usage, usageFromResponse(contentType, Buffer.concat(chunks, total)));
  } catch (error) {
    try { await reader.cancel(); } catch {}
    if (!response.destroyed) response.destroy(error);
  }
}

async function createLoopbackGateway({ baseUrl, apiKey, model, allowedTools = [], fetchImpl = globalThis.fetch }) {
  if (typeof baseUrl !== "string" || !baseUrl.startsWith("https://") && !baseUrl.startsWith("http://127.0.0.1:")) {
    throw new Error("模型服务地址无效");
  }
  if (typeof apiKey !== "string" || !apiKey || typeof model !== "string" || !model || typeof fetchImpl !== "function") {
    throw new Error("模型运行参数无效");
  }

  const key = randomBytes(35).toString("base64url");
  const state = { calls: 0, usage: {}, allowedTools: new Set(allowedTools) };
  const upstreamUrl = `${baseUrl.replace(/\/+$/, "")}/chat/completions`;
  const server = http.createServer(async (request, response) => {
    if (request.method !== "POST" || request.url !== "/v1/chat/completions" ||
        !bearerMatches(request.headers.authorization, key)) {
      replyError(response, 403, "请求未获授权");
      return;
    }

    let params;
    try {
      const raw = await readBody(request);
      params = sanitizeRequest(JSON.parse(raw.toString("utf8")), model, state.allowedTools);
    } catch (error) {
      replyError(response, error.statusCode || 422, "请求格式或工具权限无效");
      return;
    }

    state.calls += 1;
    if (state.calls > MAX_CALLS) {
      replyError(response, 429, "本轮模型调用次数已达上限");
      return;
    }

    try {
      const upstream = await fetchImpl(upstreamUrl, {
        method: "POST",
        headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json", Accept: "text/event-stream, application/json" },
        body: JSON.stringify(params),
        signal: AbortSignal.timeout(80_000)
      });
      await pipeResponse(upstream, response, state);
    } catch {
      replyError(response, 502, "模型服务连接失败");
    }
  });

  server.headersTimeout = 10_000;
  server.requestTimeout = 85_000;
  await new Promise((resolve, reject) => {
    const onError = (error) => { server.off("listening", onListening); reject(error); };
    const onListening = () => { server.off("error", onError); resolve(); };
    server.once("error", onError);
    server.once("listening", onListening);
    server.listen(0, "127.0.0.1");
  });

  const address = server.address();
  return {
    key,
    url: `http://127.0.0.1:${address.port}/v1`,
    get usage() { return { ...state.usage }; },
    async close() {
      if (!server.listening) return;
      await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    }
  };
}

module.exports = {
  MAX_CALLS,
  REQUEST_LIMIT,
  RESPONSE_LIMIT,
  TOKEN_LIMIT,
  createLoopbackGateway,
  sanitizeRequest
};
