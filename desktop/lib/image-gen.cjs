const { quotaExhaustionMessage } = require("./usage.cjs");

const MAX_IMAGE_BYTES = 10 * 1024 * 1024;
const SIZES = new Set(["1024x1024", "1536x1024", "1024x1536", "1536x864"]);
const QUALITIES = new Set(["low", "medium", "high"]);
const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const USAGE_FIELDS = ["input_tokens", "output_tokens", "total_tokens"];
const DETAIL_FIELDS = ["image_tokens", "text_tokens"];

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function safeInteger(value) {
  return Number.isSafeInteger(value) && value >= 0;
}

function cleanUsage(value) {
  if (!isObject(value)) return {};
  const result = {};
  for (const field of USAGE_FIELDS) {
    if (safeInteger(value[field])) result[field] = value[field];
  }
  for (const field of ["input_tokens_details", "output_tokens_details"]) {
    if (!isObject(value[field])) continue;
    const details = Object.fromEntries(DETAIL_FIELDS
      .filter((name) => safeInteger(value[field][name]))
      .map((name) => [name, value[field][name]]));
    if (Object.keys(details).length) result[field] = details;
  }
  return result;
}

function imageBuffer(value) {
  if (typeof value !== "string" || !value || value.length % 4 !== 0 ||
      value.length > Math.ceil(MAX_IMAGE_BYTES / 3) * 4) {
    throw new Error("生图返回内容无效");
  }
  const bytes = Buffer.from(value, "base64");
  if (bytes.toString("base64") !== value || bytes.length > MAX_IMAGE_BYTES || bytes.length < 24 ||
      !bytes.subarray(0, 8).equals(PNG_SIGNATURE) ||
      bytes.readUInt32BE(8) !== 13 || bytes.toString("ascii", 12, 16) !== "IHDR") {
    throw new Error("生图返回内容无效");
  }
  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  if (!width || !height) throw new Error("生图返回内容无效");
  return { bytes, width, height };
}

function upstreamMessage(payload) {
  const message = payload?.error?.message;
  return typeof message === "string" ? message : "";
}

function mappedError(status, payload, apiKey, now) {
  const message = upstreamMessage(payload).replaceAll(apiKey, "[已隐藏]");
  if ([402, 403].includes(status) && /balance|余额/i.test(message)) {
    return "账户余额不足，无法生图";
  }
  if (status === 403 && payload?.error?.type === "permission_error") {
    return "你所在分组未开放生图";
  }
  if (status === 429) {
    const error = payload?.error || {};
    const detail = {
      type: typeof error.type === "string" ? error.type : "",
      code: typeof error.code === "string" ? error.code : "",
      message,
      resetAt: error.metadata?.window_resets_at,
      retryAfter: Number(payload?.retry_after)
    };
    return (quotaExhaustionMessage(detail, undefined, { now: now() }) || "模型额度已用尽").replaceAll(apiKey, "[已隐藏]");
  }
  if (status === 400) return `生图请求被拒绝：${(message || "请求内容不符合要求").slice(0, 200)}`;
  return `生图失败（HTTP ${status}）`;
}

function createImageGenerator({ getModels, getCredential, fetchImpl, now = Date.now, timeoutMs = 180_000 }) {
  if (typeof getModels !== "function" || typeof getCredential !== "function" || typeof fetchImpl !== "function" ||
      typeof now !== "function" || !Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    throw new TypeError("生图运行参数无效");
  }
  let running = false;

  return async function generateImage(request, callOptions = {}) {
    if (running) throw new Error("上一张图片还在生成");
    running = true;
    try {
      if (!isObject(request)) throw new Error("生图请求无效");
      const allowedFields = new Set(["model", "prompt", "size", "quality"]);
      if (Object.keys(request).some((field) => !allowedFields.has(field))) throw new Error("生图请求无效");
      const model = request.model;
      const prompt = typeof request.prompt === "string" ? request.prompt.trim() : "";
      const size = request.size;
      const quality = request.quality;
      if (typeof model !== "string" || !/^gpt-image-/.test(model) || typeof prompt !== "string" ||
          prompt.length < 1 || prompt.length > 4000 || !SIZES.has(size) || !QUALITIES.has(quality)) {
        throw new Error("生图请求无效");
      }
      const timeout = callOptions.timeoutMs ?? timeoutMs;
      if (!Number.isFinite(timeout) || timeout <= 0) throw new Error("生图请求无效");
      const models = await getModels();
      if (!Array.isArray(models) || !models.includes(model)) throw new Error("该生图模型未对你开放");
      const credential = await getCredential();
      if (typeof credential?.baseUrl !== "string" || typeof credential?.apiKey !== "string" || !credential.apiKey) {
        throw new Error("模型凭证无效，请重新登录后重试");
      }
      const baseUrl = credential.baseUrl.replace(/\/+$/, "");
      if (!/^https:\/\//.test(baseUrl) && !/^http:\/\/127\.0\.0\.1(?::\d+)?\/v1$/.test(baseUrl)) {
        throw new Error("模型凭证返回的服务地址无效");
      }
      const controller = new AbortController();
      let timedOut = false;
      const timer = setTimeout(() => {
        timedOut = true;
        controller.abort();
      }, timeout);
      const abortFromCaller = () => controller.abort(callOptions.signal?.reason);
      if (callOptions.signal?.aborted) abortFromCaller();
      else callOptions.signal?.addEventListener("abort", abortFromCaller, { once: true });
      try {
        const response = await fetchImpl(`${baseUrl}/images/generations`, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${credential.apiKey}`,
            "Content-Type": "application/json",
            Accept: "application/json"
          },
          body: JSON.stringify({ model, prompt, size, quality, n: 1 }),
          redirect: "error",
          signal: controller.signal
        });
        let payload;
        try { payload = await response.json(); } catch {
          if (!response.ok) throw new Error(`生图失败（HTTP ${response.status}）`);
          throw new Error("生图返回内容无效");
        }
        if (!response.ok) throw new Error(mappedError(response.status, payload, credential.apiKey, now));
        const encoded = payload?.data?.[0]?.b64_json;
        const { bytes, width, height } = imageBuffer(encoded);
        const [expectedWidth, expectedHeight] = size.split("x").map(Number);
        if (width !== expectedWidth || height !== expectedHeight) throw new Error("生图返回内容无效");
        return {
          png_base64: bytes.toString("base64"), width, height, model, size, quality, prompt,
          usage: cleanUsage(payload.usage)
        };
      } catch (error) {
        if (timedOut) {
          throw new Error("生图超时，请稍后重试（如已扣费以 sub2api 记录为准）");
        }
        throw error;
      } finally {
        clearTimeout(timer);
        callOptions.signal?.removeEventListener("abort", abortFromCaller);
      }
    } finally {
      running = false;
    }
  };
}

module.exports = { MAX_IMAGE_BYTES, SIZES, QUALITIES, imageBuffer, createImageGenerator };
