const test = require("node:test");
const assert = require("node:assert/strict");
const { MAX_IMAGE_BYTES, createImageGenerator } = require("../lib/image-gen.cjs");

const MODEL = "gpt-image-2.6-flare";
const credential = { baseUrl: "https://images.example/v1", apiKey: "fakekey-image-test" };

function png(width = 1024, height = 1024, size = 24) {
  const bytes = Buffer.alloc(size);
  Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]).copy(bytes);
  bytes.writeUInt32BE(13, 8);
  bytes.write("IHDR", 12, 4, "ascii");
  bytes.writeUInt32BE(width, 16);
  bytes.writeUInt32BE(height, 20);
  return bytes;
}

function response(status, payload) {
  return { ok: status >= 200 && status < 300, status, json: async () => payload };
}

function generator({ fetchImpl, models = [MODEL], timeoutMs = 180_000, now = () => 1_000 }) {
  return createImageGenerator({ getModels: async () => models, getCredential: async () => credential,
    fetchImpl, timeoutMs, now });
}

test("图像白名单拒绝不可用模型和非法参数且不调用 fetch", async () => {
  const calls = [];
  const generate = generator({ fetchImpl: async (...args) => { calls.push(args); return response(200, {}); } });
  const invalid = [
    [{ model: "gpt-image-not-open", prompt: "x", size: "1024x1024", quality: "low" }, /未对你开放/],
    [{ model: "chat-model", prompt: "x", size: "1024x1024", quality: "low" }, /无效/],
    [{ model: MODEL, prompt: "x", size: "512x512", quality: "low" }, /无效/],
    [{ model: MODEL, prompt: "x", size: "1024x1024", quality: "ultra" }, /无效/],
    [{ model: MODEL, prompt: "   ", size: "1024x1024", quality: "low" }, /无效/],
    [{ model: MODEL, prompt: "x".repeat(4001), size: "1024x1024", quality: "low" }, /无效/],
    [{ model: MODEL, prompt: "x", size: "1024x1024", quality: "low", response_format: "url" }, /无效/]
  ];
  for (const [request, error] of invalid) await assert.rejects(generate(request), error);
  assert.equal(calls.length, 0);
});

test("成功请求只发送白名单字段并按 IHDR 读取尺寸与 usage", async () => {
  const calls = [];
  const encoded = png().toString("base64");
  const generate = generator({ fetchImpl: async (url, options) => {
    calls.push({ url, options });
    return response(200, { data: [{ b64_json: encoded }], usage: {
      input_tokens: 7, output_tokens: 9, total_tokens: 16,
      input_tokens_details: { image_tokens: 3, text_tokens: 4, secret: 100 },
      output_tokens_details: { image_tokens: 8, text_tokens: 1 }, secret: "drop"
    } });
  } });
  const result = await generate({ model: MODEL, prompt: "  a test image  ", size: "1024x1024", quality: "medium" });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, "https://images.example/v1/images/generations");
  assert.equal(calls[0].options.headers.Authorization, `Bearer ${credential.apiKey}`);
  assert.equal(calls[0].options.redirect, "error");
  const body = JSON.parse(calls[0].options.body);
  assert.deepEqual(Object.keys(body).sort(), ["model", "n", "prompt", "quality", "size"]);
  assert.equal(body.n, 1);
  assert.equal(body.prompt, "a test image");
  assert.equal(result.prompt, "a test image");
  assert.equal(result.width, 1024);
  assert.equal(result.height, 1024);
  assert.deepEqual(result.usage, { input_tokens: 7, output_tokens: 9, total_tokens: 16,
    input_tokens_details: { image_tokens: 3, text_tokens: 4 },
    output_tokens_details: { image_tokens: 8, text_tokens: 1 } });
  assert.equal(JSON.stringify(result).includes(credential.apiKey), false);
});

test("上游状态映射区分分组、余额、额度、审核和其他错误", async (t) => {
  const scenarios = [
    [403, { error: { type: "permission_error", message: "Image generation is not enabled" } }, "你所在分组未开放生图"],
    [402, { error: { message: "balance too low" } }, "账户余额不足，无法生图"],
    [403, { error: { type: "permission_error", message: "balance insufficient" } }, "账户余额不足，无法生图"],
    [429, { error: { type: "rate_limit_exceeded", code: "quota_exhausted", message: "quota exhausted" } }, /额度/],
    [400, { error: { message: `review ${credential.apiKey} denied` } }, /生图请求被拒绝：review \[已隐藏\] denied/],
    [500, { error: { message: "internal detail" } }, "生图失败（HTTP 500）"]
  ];
  for (const [status, payload, expected] of scenarios) await t.test(`${status} ${String(expected)}`, async () => {
    const generate = generator({ fetchImpl: async () => response(status, payload) });
    const matches = expected instanceof RegExp ? expected : (error) => error.message === expected;
    await assert.rejects(generate({ model: MODEL, prompt: "x", size: "1024x1024", quality: "low" }), matches);
  });
});

test("无效 PNG、缺图、越限和尺寸不符均拒绝", async (t) => {
  const tooLarge = png(1024, 1024, MAX_IMAGE_BYTES + 1).toString("base64");
  const cases = [
    ["非 PNG", { data: [{ b64_json: Buffer.from("not png").toString("base64") }] }],
    ["缺少 b64_json", { data: [{}] }],
    ["大于 10 MB", { data: [{ b64_json: tooLarge }] }],
    ["IHDR 尺寸与请求不符", { data: [{ b64_json: png(1536, 1024).toString("base64") }] }]
  ];
  for (const [name, payload] of cases) await t.test(name, async () => {
    const generate = generator({ fetchImpl: async () => response(200, payload) });
    await assert.rejects(generate({ model: MODEL, prompt: "x", size: "1024x1024", quality: "low" }),
      /生图返回内容无效/);
  });
});

test("超时会真实触发 AbortSignal，按钮并发请求直接拒绝", async () => {
  const timeout = generator({ timeoutMs: 15, fetchImpl: (_url, options) => new Promise((_resolve, reject) => {
    options.signal.addEventListener("abort", () => reject(Object.assign(new Error("aborted"), { name: "AbortError" })), { once: true });
  }) });
  await assert.rejects(timeout({ model: MODEL, prompt: "x", size: "1024x1024", quality: "low" }), /生图超时/);

  let release;
  let started;
  const inFetch = new Promise((resolve) => { started = resolve; });
  const pending = new Promise((resolve) => { release = resolve; });
  const concurrent = generator({ fetchImpl: async () => { started(); await pending; return response(200, { data: [{ b64_json: png().toString("base64") }] }); } });
  const first = concurrent({ model: MODEL, prompt: "x", size: "1024x1024", quality: "low" });
  await inFetch;
  await assert.rejects(concurrent({ model: MODEL, prompt: "y", size: "1024x1024", quality: "low" }), /上一张图片还在生成/);
  release();
  const result = await first;
  assert.equal(result.width, 1024);
});
