const test = require("node:test");
const assert = require("node:assert/strict");
const { createImageBroker } = require("../lib/image-broker.cjs");

const model = "gpt-image-2.6-flare";
const options = { prompt: "a desk lamp", size: "1536x1024", quality: "medium" };
const generatedImage = { png_base64: "secret-base64-image", width: 1536, height: 1024, model,
  size: options.size, quality: options.quality, prompt: options.prompt, usage: { output_tokens: 4 } };

async function send(broker, body = options, { key = broker.key, path = "/generate", method = "POST", raw } = {}) {
  const response = await fetch(broker.url + path, {
    method,
    headers: { Authorization: `Bearer ${key}`, ...(raw === undefined ? { "Content-Type": "application/json" } : {}) },
    ...(method === "POST" ? { body: raw ?? JSON.stringify(body) } : {})
  });
  const text = await response.text();
  assert.equal(text.includes(broker.key), false, "broker key must not be returned");
  assert.equal(text.includes("secret-base64-image"), false, "PNG payload must not be returned");
  return { status: response.status, value: JSON.parse(text) };
}

test("broker 对错误 Bearer、路径、方法、超大和非法 body 作出拒绝且不调用 generate", async () => {
  let calls = 0;
  const broker = await createImageBroker({ generate: async () => { calls += 1; return generatedImage; }, model,
    deadlineAt: Date.now() + 300_000 });
  try {
    const denied = await send(broker, options, { key: "wrong-key" });
    assert.equal(denied.status, 401);
    assert.equal(denied.value.ok, false);
    const path = await send(broker, options, { path: "/other" });
    assert.equal(path.status, 404);
    const method = await send(broker, options, { method: "GET" });
    assert.equal(method.status, 404);
    const large = await send(broker, null, { raw: JSON.stringify({ prompt: "x".repeat(17_000) }) });
    assert.equal(large.status, 413);
    const invalids = [
      { prompt: "  ", size: options.size, quality: options.quality },
      { ...options, model },
      { ...options, quality: "unknown" },
      { ...options, size: "512x512" }
    ];
    for (const value of invalids) {
      const result = await send(broker, value);
      assert.equal(result.status, 400);
      assert.equal(result.value.ok, false);
    }
    assert.equal(calls, 0);
  } finally {
    await broker.close();
  }
});

test("尝试次数占位先于 generate，失败也计入两张上限且 images 逐字段保存", async () => {
  let calls = 0;
  const broker = await createImageBroker({
    generate: async () => {
      calls += 1;
      if (calls === 2) throw new Error("上游失败但可能已扣费");
      return generatedImage;
    }, model, maxImages: 2, deadlineAt: Date.now() + 300_000
  });
  try {
    const first = await send(broker);
    assert.equal(first.status, 200);
    assert.deepEqual(first.value, { ok: true, index: 1, width: 1536, height: 1024,
      size: "1536x1024", quality: "medium" });
    const second = await send(broker);
    assert.equal(second.status, 502);
    assert.equal(second.value.ok, false);
    assert.match(second.value.message, /上游失败/);
    const third = await send(broker);
    assert.equal(third.status, 409);
    assert.equal(third.value.message, "本轮生图次数已达上限（2 张）");
    assert.equal(calls, 2);
    assert.equal(broker.images.length, 1);
    assert.deepEqual(broker.images[0], generatedImage);
  } finally {
    await broker.close();
  }
});

test("第二个并发生图被拒绝且剩余预算不足时不调用 generate", async () => {
  let release;
  let started;
  let calls = 0;
  const firstStarted = new Promise((resolve) => { started = resolve; });
  const pending = new Promise((resolve) => { release = resolve; });
  const broker = await createImageBroker({
    generate: async () => { calls += 1; started(); await pending; return generatedImage; }, model,
    deadlineAt: Date.now() + 300_000
  });
  try {
    const firstPromise = send(broker);
    await firstStarted;
    const concurrent = await send(broker);
    assert.equal(concurrent.status, 429);
    assert.equal(concurrent.value.message, "上一张图片还在生成");
    release();
    const first = await firstPromise;
    assert.equal(first.status, 200);
    assert.equal(calls, 1);
  } finally {
    release?.();
    await broker.close();
  }

  let deniedCalls = 0;
  const now = 10_000;
  const short = await createImageBroker({ generate: async () => { deniedCalls += 1; return generatedImage; },
    model, deadlineAt: now + 64_999, now: () => now });
  try {
    const denied = await send(short);
    assert.equal(denied.status, 409);
    assert.equal(denied.value.message, "本轮剩余时间不足，未生成图片");
    assert.equal(deniedCalls, 0);
  } finally {
    await short.close();
  }
});
