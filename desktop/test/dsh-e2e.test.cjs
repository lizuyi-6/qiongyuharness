const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { createLoopbackGateway } = require("../lib/loopback-gateway.cjs");
const { runDsh } = require("../lib/dsh-bridge.cjs");
const { buildRuntimeEnv } = require("../lib/runtime-env.cjs");
const { readJson, sseResponse, startServer } = require("./helpers.cjs");

const desktopDirectory = path.resolve(__dirname, "..");
const repositoryDirectory = path.resolve(desktopDirectory, "..");
const harnessDirectory = path.join(repositoryDirectory, "harness");
const dshBin = path.join(harnessDirectory, "node_modules", "@deepseek-ai", "dsh", "lib", "bin.js");
const patchPath = path.join(harnessDirectory, "office.patch.yml");
const MODEL = "desktop-e2e-model";

test("真实 Harness DSH 经 JS 桥连接假 SSE 网关并返回文本与 usage", { timeout: 110_000 }, async (t) => {
  await fs.access(dshBin);
  await fs.access(patchPath);
  assert.match(process.versions.node, /^24\./);

  const captured = [];
  const upstream = await startServer(async (request, response) => {
    captured.push({ headers: request.headers, body: await readJson(request) });
    sseResponse(response, {
      text: "JS 桥接真实 DSH 的验收响应",
      usage: { prompt_tokens: 50, completion_tokens: 40, total_tokens: 90 }
    });
  });
  const place = await fs.mkdtemp(path.join(os.tmpdir(), "qo-desktop-dsh-test-"));
  let gateway;
  try {
    gateway = await createLoopbackGateway({
      baseUrl: upstream.url,
      apiKey: "test-only-upstream-key",
      model: MODEL
    });
    const env = buildRuntimeEnv({
      baseEnv: process.env,
      home: place,
      gatewayUrl: gateway.url,
      gatewayKey: gateway.key,
      model: MODEL,
      system: "只按用户要求生成简洁文书。"
    });
    assert.notEqual(env.QO_RUNTIME_KEY, "test-only-upstream-key");
    const text = await runDsh({
      executable: process.execPath,
      dshBin,
      patchPath,
      cwd: place,
      env,
      model: MODEL,
      prompt: "请输出验收文本。"
    });
    assert.equal(text, "JS 桥接真实 DSH 的验收响应");
    assert.deepEqual(gateway.usage, { prompt_tokens: 50, completion_tokens: 40, total_tokens: 90 });
    assert.equal(captured.length, 1);
    assert.equal(captured[0].headers.authorization, "Bearer test-only-upstream-key");
    assert.equal(captured[0].body.model, MODEL);
    assert.equal(captured[0].body.stream, true);
  } finally {
    if (gateway) await gateway.close();
    await upstream.close();
    await fs.rm(place, { recursive: true, force: true });
  }
});
