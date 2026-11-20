const { createServer } = require("node:http");
const { randomBytes, timingSafeEqual } = require("node:crypto");
const { QUALITIES, SIZES } = require("./image-gen.cjs");

const MAX_BODY_BYTES = 16 * 1024;
const MAX_IMAGE_TIMEOUT_MS = 120_000;
const MIN_IMAGE_TIMEOUT_MS = 45_000;

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

function validOptions(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  if (Object.keys(value).some((key) => !["prompt", "size", "quality"].includes(key))) return null;
  const prompt = typeof value.prompt === "string" ? value.prompt.trim() : "";
  if (prompt.length < 1 || prompt.length > 4000 || !SIZES.has(value.size) || !QUALITIES.has(value.quality)) return null;
  return { prompt, size: value.size, quality: value.quality };
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

async function createImageBroker({ generate, model, maxImages = 2, deadlineAt, now = Date.now }) {
  if (typeof generate !== "function" || typeof model !== "string" || !/^gpt-image-/.test(model) ||
      !Number.isInteger(maxImages) || maxImages < 1 || !Number.isFinite(deadlineAt) || typeof now !== "function") {
    throw new TypeError("生图 broker 参数无效");
  }
  const key = randomBytes(32).toString("base64url");
  const secret = Buffer.from(key, "utf8");
  const images = [];
  let attempts = 0;
  let busy = false;

  const server = createServer(async (request, response) => {
    if (request.method !== "POST" || request.url !== "/generate") {
      json(response, 404, { ok: false, message: "未找到请求路径" });
      request.resume();
      return;
    }
    const authorization = request.headers.authorization || "";
    const bearer = authorization.startsWith("Bearer ") ? authorization.slice(7) : "";
    if (!equalSecret(bearer, secret)) {
      json(response, 401, { ok: false, message: "生图工具认证失败" });
      request.resume();
      return;
    }
    const body = await readBody(request);
    if (body.tooLarge) {
      json(response, 413, { ok: false, message: "生图请求超过大小限制" });
      return;
    }
    if (body.invalid) {
      json(response, 400, { ok: false, message: "生图请求格式无效" });
      return;
    }
    const options = validOptions(body.value);
    if (!options) {
      json(response, 400, { ok: false, message: "生图请求参数无效" });
      return;
    }
    if (busy) {
      json(response, 429, { ok: false, message: "上一张图片还在生成" });
      return;
    }
    if (attempts >= maxImages) {
      json(response, 409, { ok: false, message: `本轮生图次数已达上限（${maxImages} 张）` });
      return;
    }
    const timeoutMs = Math.min(MAX_IMAGE_TIMEOUT_MS, deadlineAt - now() - 20_000);
    if (timeoutMs < MIN_IMAGE_TIMEOUT_MS) {
      json(response, 409, { ok: false, message: "本轮剩余时间不足，未生成图片" });
      return;
    }

    attempts += 1;
    busy = true;
    const controller = new AbortController();
    request.once("aborted", () => controller.abort(new Error("生图请求已取消")));
    response.once("close", () => {
      if (!response.writableEnded) controller.abort(new Error("生图请求已取消"));
    });
    try {
      const image = await generate({ ...options, model }, { timeoutMs, signal: controller.signal });
      if (!image || image.model !== model || !Number.isSafeInteger(image.width) || image.width < 1 ||
          !Number.isSafeInteger(image.height) || image.height < 1 || !SIZES.has(image.size) ||
          !QUALITIES.has(image.quality) || typeof image.png_base64 !== "string") {
        throw new Error("生图返回内容无效");
      }
      const index = images.length + 1;
      images.push({ ...image });
      json(response, 200, { ok: true, index, width: image.width, height: image.height,
        size: image.size, quality: image.quality });
    } catch (error) {
      const raw = typeof error?.message === "string" ? error.message : "生图失败";
      const message = raw.replaceAll(key, "[已隐藏]").replace(/[A-Za-z0-9+/]{32,}={0,2}/g, "[内容已隐藏]").slice(0, 300);
      json(response, 502, { ok: false, message: message || "生图失败" });
    } finally {
      busy = false;
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
    images,
    async close() {
      if (!server.listening) return;
      await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    }
  };
}

module.exports = { MAX_BODY_BYTES, MAX_IMAGE_TIMEOUT_MS, MIN_IMAGE_TIMEOUT_MS, createImageBroker };
