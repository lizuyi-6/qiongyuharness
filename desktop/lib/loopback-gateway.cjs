const http = require("node:http");
const { randomBytes, timingSafeEqual } = require("node:crypto");
const { StringDecoder } = require("node:string_decoder");
const { isIsoDateTime } = require("./usage.cjs");

const REQUEST_LIMIT = 2 * 1024 * 1024;
const RESPONSE_LIMIT = 3 * 1024 * 1024;
const MAX_CALLS = 32;
// 不再自设输出上限：表内模型按其官方单次输出上限下发 max_tokens（2026-11-20 查官方/公开文档，
// 并在线上以该值实测上游返回 200）；表外模型不下发 max_tokens，由上游默认值决定。
// 显式下发而不是省略，是因为 sub2api 转 Anthropic 协议上游时，缺省会补 8192。
const MODEL_OUTPUT_LIMITS = Object.freeze({
  "glm-5.3": 131_072,
  "step-5-preview": 64_000,
  "deepseek-v4.1-flash-1017": 393_216,
  "k3.1-pro": 131_072
});
// 推理强度一律用各模型的最高档（公司要求质量优先，2026-11-20）。依据官方/公开文档：
// glm-5.3 low/high/max（默认 max）；k3.1-pro low/high/max（默认 max）；
// deepseek-v4.1-flash low/high/max（默认 high）；step-5-preview 公开文档最高为 high。
// sub2api 把 step / deepseek 转成 Anthropic 协议时映射为 output_config.effort + thinking。
const MODEL_REASONING_EFFORT = Object.freeze({
  "glm-5.3": "max",
  "step-5-preview": "high",
  "deepseek-v4.1-flash-1017": "max",
  "k3.1-pro": "max"
});
// 单次上游调用的时限。最高推理强度下执行阶段实测 95 秒（默认强度），留足余量。
const UPSTREAM_TIMEOUT_MS = 600_000;
const FORWARDED_FIELDS = [
  "model", "messages", "stream", "stream_options", "temperature", "top_p",
  "reasoning_effort", "thinking"
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

  // DSH 请求里的 max_tokens 来自它的内部预算配置，不转发。
  if (Object.hasOwn(MODEL_OUTPUT_LIMITS, model)) params.max_tokens = MODEL_OUTPUT_LIMITS[model];
  if (Object.hasOwn(MODEL_REASONING_EFFORT, model)) params.reasoning_effort = MODEL_REASONING_EFFORT[model];
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

function responseDetails(contentType, body) {
  const finishReasons = [];
  let usage = {};
  const collect = (payload) => {
    const nextUsage = usageFromPayload(payload);
    if (Object.keys(nextUsage).length) usage = nextUsage;
    for (const choice of Array.isArray(payload?.choices) ? payload.choices : []) {
      if (typeof choice?.finish_reason === "string") finishReasons.push(choice.finish_reason);
    }
  };
  if (contentType.includes("text/event-stream")) {
    for (const line of body.toString("utf8").split(/\r?\n/)) {
      if (!line.startsWith("data:")) continue;
      const data = line.slice(5).trim();
      if (!data || data === "[DONE]") continue;
      try {
        collect(JSON.parse(data));
      } catch {}
    }
    return { usage, finishReasons };
  }
  try {
    collect(JSON.parse(body.toString("utf8")));
  } catch {
    return { usage: {}, finishReasons: [] };
  }
  return { usage, finishReasons };
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

async function readErrorBody(upstream, limit = 64 * 1024) {
  if (!upstream.body) return "";
  const reader = upstream.body.getReader();
  const chunks = [];
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > limit) {
        await reader.cancel();
        return "";
      }
      chunks.push(Buffer.from(value));
    }
    return Buffer.concat(chunks, total).toString("utf8");
  } catch {
    try { await reader.cancel(); } catch {}
    return "";
  }
}

function quotaErrorDetails(upstream, body) {
  if (upstream.status !== 429) return null;
  let payload;
  try { payload = JSON.parse(body); } catch { return null; }
  const error = payload?.error;
  if (!isObject(error) || error.type !== "rate_limit_exceeded" && error.code !== "rate_limit_exceeded") return null;

  const details = { type: "rate_limit_exceeded" };
  if (typeof error.code === "string" && error.code.length <= 120) details.code = error.code;
  if (typeof error.message === "string" && error.message.length <= 400) details.message = error.message;
  const resetAt = error.metadata?.window_resets_at;
  if (isIsoDateTime(resetAt)) details.resetAt = resetAt;
  const retryAfter = Number(upstream.headers.get("retry-after"));
  if (Number.isFinite(retryAfter) && retryAfter > 0 && retryAfter <= 30 * 24 * 60 * 60) details.retryAfter = retryAfter;
  return details;
}

function extractDeltaSignals(payload, seenTools, onStreamEvent) {
  if (!payload || typeof payload !== "object") return;
  for (const choice of Array.isArray(payload.choices) ? payload.choices : []) {
    const delta = choice?.delta || choice?.message;
    if (!delta || typeof delta !== "object") continue;
    const thinking = typeof delta.reasoning_content === "string"
      ? delta.reasoning_content
      : typeof delta.reasoning === "string"
        ? delta.reasoning
        : typeof delta.thinking === "string"
          ? delta.thinking
          : "";
    if (thinking) onStreamEvent({ type: "thinking", text: thinking });
    if (Array.isArray(delta.tool_calls)) {
      for (const call of delta.tool_calls) {
        const name = call?.function?.name;
        const idx = call?.index ?? call?.id ?? name;
        const key = `${idx}:${name}`;
        if (typeof name === "string" && name && !seenTools.has(key)) {
          seenTools.add(key);
          onStreamEvent({ type: "tool", name });
        }
      }
    }
    if (typeof delta.content === "string" && delta.content) {
      onStreamEvent({ type: "content", chars: delta.content.length });
    }
  }
}

async function pipeResponse(upstream, response, state) {
  const contentType = upstream.headers.get("content-type") || "application/json";
  if (upstream.ok && !state.sticky_quota_error) state.quota_error = null;
  if (!upstream.ok) {
    const body = await readErrorBody(upstream);
    const quotaError = quotaErrorDetails(upstream, body);
    if (quotaError) {
      state.quota_error = quotaError;
      replyError(response, 429, "模型额度已用尽");
      return;
    }
    replyError(response, 502, "模型服务请求失败");
    return;
  }
  if (!upstream.body) {
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
  const isSse = contentType.includes("text/event-stream");
  const decoder = isSse ? new StringDecoder("utf8") : null;
  const seenTools = new Set();
  let sseBuffer = "";
  const parseSseLine = (line) => {
    if (!line.startsWith("data:")) return;
    const raw = line.slice(5).trim();
    if (!raw || raw === "[DONE]") return;
    try {
      extractDeltaSignals(JSON.parse(raw), seenTools, state.onStreamEvent);
    } catch {}
  };
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
      if (isSse && typeof state.onStreamEvent === "function") {
        sseBuffer += decoder.write(chunk);
        let nl;
        while ((nl = sseBuffer.indexOf("\n")) >= 0) {
          const line = sseBuffer.slice(0, nl).replace(/\r$/, "");
          sseBuffer = sseBuffer.slice(nl + 1);
          parseSseLine(line);
        }
      }
      if (!response.write(chunk)) await new Promise((resolve) => response.once("drain", resolve));
    }
    if (isSse && typeof state.onStreamEvent === "function") {
      sseBuffer += decoder.end();
      if (sseBuffer.trim()) parseSseLine(sseBuffer.replace(/\r$/, ""));
    }
    response.end();
    const fullBody = Buffer.concat(chunks, total);
    if (!isSse && typeof state.onStreamEvent === "function") {
      try {
        extractDeltaSignals(JSON.parse(fullBody.toString("utf8")), seenTools, state.onStreamEvent);
      } catch {}
    }
    const details = responseDetails(contentType, fullBody);
    addUsage(state.usage, details.usage);
    for (const reason of details.finishReasons) {
      if (state.finish_reasons.length < MAX_CALLS) state.finish_reasons.push(reason);
      if (reason === "length") state.truncated = true;
    }
  } catch (error) {
    try { await reader.cancel(); } catch {}
    if (!response.destroyed) response.destroy(error);
  }
}

async function createLoopbackGateway({ baseUrl, apiKey, model, allowedTools = [], stickyQuotaError = false,
  fetchImpl = globalThis.fetch, onStreamEvent = null }) {
  if (typeof baseUrl !== "string" || !baseUrl.startsWith("https://") && !baseUrl.startsWith("http://127.0.0.1:")) {
    throw new Error("模型服务地址无效");
  }
  if (typeof apiKey !== "string" || !apiKey || typeof model !== "string" || !model || typeof fetchImpl !== "function") {
    throw new Error("模型运行参数无效");
  }

  const key = randomBytes(35).toString("base64url");
  const state = {
    calls: 0,
    usage: {},
    allowedTools: new Set(allowedTools),
    max_tokens_clamped: [],
    finish_reasons: [],
    truncated: false,
    quota_error: null,
    sticky_quota_error: stickyQuotaError === true,
    onStreamEvent: typeof onStreamEvent === "function" ? onStreamEvent : null
  };
  const upstreamUrl = `${baseUrl.replace(/\/+$/, "")}/chat/completions`;
  const server = http.createServer(async (request, response) => {
    if (request.method !== "POST" || request.url !== "/v1/chat/completions" ||
        !bearerMatches(request.headers.authorization, key)) {
      replyError(response, 403, "请求未获授权");
      return;
    }
    if (state.sticky_quota_error && state.quota_error) {
      replyError(response, 429, "模型额度已用尽");
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
        signal: AbortSignal.timeout(UPSTREAM_TIMEOUT_MS)
      });
      await pipeResponse(upstream, response, state);
    } catch {
      replyError(response, 502, "模型服务连接失败");
    }
  });

  server.headersTimeout = 10_000;
  server.requestTimeout = UPSTREAM_TIMEOUT_MS + 10_000;
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
    get max_tokens_clamped() { return state.max_tokens_clamped.map((item) => ({ ...item })); },
    get finish_reasons() { return [...state.finish_reasons]; },
    get truncated() { return state.truncated; },
    get quota_error() { return state.quota_error ? { ...state.quota_error } : null; },
    async close() {
      if (!server.listening) return;
      await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    }
  };
}

module.exports = {
  MAX_CALLS,
  MODEL_OUTPUT_LIMITS,
  MODEL_REASONING_EFFORT,
  REQUEST_LIMIT,
  RESPONSE_LIMIT,
  UPSTREAM_TIMEOUT_MS,
  createLoopbackGateway,
  sanitizeRequest
};
