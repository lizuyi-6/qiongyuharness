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

test("回环网关把 max_tokens 夹到 3500 并累计 SSE usage", async (t) => {
  let forwarded;
  const upstream = await startServer(async (request, response) => {
    forwarded = await readJson(request);
    sseResponse(response, { usage: { prompt_tokens: 4, completion_tokens: 5, total_tokens: 9 } });
  });
  const gateway = await createLoopbackGateway({ baseUrl: upstream.url, apiKey: "only-upstream", model: MODEL });
  t.after(async () => { await gateway.close(); await upstream.close(); });

  const response = await post(gateway, validRequest({ max_tokens: 9000 }));
  assert.equal(response.status, 200);
  assert.equal(forwarded.max_tokens, 3500);
  assert.deepEqual(gateway.usage, { prompt_tokens: 4, completion_tokens: 5, total_tokens: 9 });
  assert.equal(await response.text().then((value) => value.includes("fake model response")), true);
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
