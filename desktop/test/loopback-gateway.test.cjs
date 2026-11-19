const test = require("node:test");
const assert = require("node:assert/strict");
const { createLoopbackGateway } = require("../lib/loopback-gateway.cjs");
const { readJson, sseResponse, startServer } = require("./helpers.cjs");

const MODEL = "mock-model";
const validRequest = (extra = {}) => ({
  model: MODEL,
  stream: true,
  messages: [{ role: "user", content: "请整理材料" }],
  ...extra
});

async function post(gateway, body, token = gateway.key) {
  return fetch(`${gateway.url}/chat/completions`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: typeof body === "string" ? body : JSON.stringify(body)
  });
}

test("回环网关拒绝未放行的工具和错误模型", async (t) => {
  const upstream = await startServer((_request, response) => sseResponse(response));
  const gateway = await createLoopbackGateway({ baseUrl: upstream.url, apiKey: "only-upstream", model: MODEL });
  t.after(async () => { await gateway.close(); await upstream.close(); });

  const disallowedTool = await post(gateway, validRequest({ tools: [
    { type: "function", function: { name: "todo_write", parameters: { type: "object" } } }
  ] }));
  assert.equal(disallowedTool.status, 422);

  const mismatchedModel = await post(gateway, validRequest({ model: "another-model" }));
  assert.equal(mismatchedModel.status, 422);

  const wrongToken = await post(gateway, validRequest(), "wrong-token");
  assert.equal(wrongToken.status, 403);
});

test("回环网关把 5000 max_tokens 夹到 3500 并记录事件与 SSE usage", async (t) => {
  let forwarded;
  const upstream = await startServer(async (request, response) => {
    forwarded = await readJson(request);
    sseResponse(response, { usage: { prompt_tokens: 4, completion_tokens: 5, total_tokens: 9 } });
  });
  const gateway = await createLoopbackGateway({ baseUrl: upstream.url, apiKey: "only-upstream", model: MODEL });
  t.after(async () => { await gateway.close(); await upstream.close(); });

  const response = await post(gateway, validRequest({ max_tokens: 5000 }));
  assert.equal(response.status, 200);
  assert.equal(forwarded.max_tokens, 3500);
  assert.deepEqual(gateway.usage, { prompt_tokens: 4, completion_tokens: 5, total_tokens: 9 });
  assert.deepEqual(gateway.max_tokens_clamped, [{ requested: 5000, applied: 3500 }]);
  assert.equal(await response.text().then((value) => value.includes("fake model response")), true);
});

test("回环网关拒绝非正 max_tokens，null 和缺省值使用 3500 且不记夹逼", async (t) => {
  const forwarded = [];
  const upstream = await startServer(async (request, response) => {
    forwarded.push(await readJson(request));
    sseResponse(response);
  });
  const gateway = await createLoopbackGateway({ baseUrl: upstream.url, apiKey: "only-upstream", model: MODEL });
  t.after(async () => { await gateway.close(); await upstream.close(); });

  for (const max_tokens of [0, -1]) {
    const rejected = await post(gateway, validRequest({ max_tokens }));
    assert.equal(rejected.status, 422);
    await rejected.text();
  }
  for (const extra of [{ max_tokens: null }, {}]) {
    const accepted = await post(gateway, validRequest(extra));
    assert.equal(accepted.status, 200);
    await accepted.text();
  }
  assert.deepEqual(forwarded.map((request) => request.max_tokens), [3500, 3500]);
  assert.deepEqual(gateway.max_tokens_clamped, []);
});

test("finish_reason 为 length 时记录 truncated，SSE 与 JSON 响应都解析", async (t) => {
  let streaming = true;
  const upstream = await startServer(async (_request, response) => {
    if (streaming) sseResponse(response, { finishReason: "length" });
    else {
      response.writeHead(200, { "Content-Type": "application/json" });
      response.end(JSON.stringify({ choices: [{ finish_reason: "length" }], usage: { total_tokens: 3 } }));
    }
    streaming = false;
  });
  const gateway = await createLoopbackGateway({ baseUrl: upstream.url, apiKey: "only-upstream", model: MODEL });
  t.after(async () => { await gateway.close(); await upstream.close(); });

  for (let count = 0; count < 2; count += 1) {
    const response = await post(gateway, validRequest({ stream: count === 0 }));
    assert.equal(response.status, 200);
    await response.text();
  }
  assert.deepEqual(gateway.finish_reasons, ["length", "length"]);
  assert.equal(gateway.truncated, true);
});

test("回环网关拒绝第五次模型调用", async (t) => {
  const upstream = await startServer((_request, response) => sseResponse(response));
  const gateway = await createLoopbackGateway({ baseUrl: upstream.url, apiKey: "only-upstream", model: MODEL });
  t.after(async () => { await gateway.close(); await upstream.close(); });

  const statuses = [];
  for (let call = 0; call < 5; call += 1) {
    const response = await post(gateway, validRequest());
    statuses.push(response.status);
    await response.text();
  }
  assert.deepEqual(statuses, [200, 200, 200, 200, 429]);
});

test("回环网关拒绝超过 2 MB 的请求体", async (t) => {
  const upstream = await startServer((_request, response) => sseResponse(response));
  const gateway = await createLoopbackGateway({ baseUrl: upstream.url, apiKey: "only-upstream", model: MODEL });
  t.after(async () => { await gateway.close(); await upstream.close(); });

  const response = await post(gateway, JSON.stringify(validRequest({ content: "x".repeat(2 * 1024 * 1024) })));
  assert.equal(response.status, 413);
});
