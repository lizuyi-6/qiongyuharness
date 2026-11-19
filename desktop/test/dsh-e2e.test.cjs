const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { createLoopbackGateway } = require("../lib/loopback-gateway.cjs");
const { runDsh } = require("../lib/dsh-bridge.cjs");
const { buildRuntimeEnv } = require("../lib/runtime-env.cjs");
const { createImageBroker } = require("../lib/image-broker.cjs");
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

test("真实 Harness DSH 调用 office_image_generate，工具结果只含短文本", { timeout: 200_000 }, async () => {
  await fs.access(dshBin);
  await fs.access(patchPath);
  assert.match(process.versions.node, /^24\./);

  const captured = [];
  const brokerCalls = [];
  const broker = await createImageBroker({
    model: "gpt-image-2.6-flare",
    deadlineAt: Date.now() + 300_000,
    generate: async (request) => {
      brokerCalls.push(request);
      return { png_base64: "private-generated-image-base64", width: 1024, height: 1024,
        model: "gpt-image-2.6-flare", size: request.size, quality: request.quality,
        prompt: request.prompt, usage: { output_tokens: 1 } };
    }
  });
  const upstream = await startServer(async (request, response) => {
    const body = await readJson(request);
    captured.push(body);
    if (captured.length === 1) {
      sseResponse(response, { toolCalls: [{ index: 0, id: "call-image-1", type: "function", function: {
        name: "office_image_generate", arguments: JSON.stringify({ prompt: "蓝色山峦的简洁封面",
          size: "1536x1024", quality: "high" })
      } }] });
    } else {
      sseResponse(response, { text: "图片已生成，作为本轮回复的配图。" });
    }
  });
  const place = await fs.mkdtemp(path.join(os.tmpdir(), "qo-desktop-image-dsh-test-"));
  let gateway;
  try {
    gateway = await createLoopbackGateway({ baseUrl: upstream.url, apiKey: "test-only-upstream-key", model: MODEL,
      allowedTools: ["office_image_generate"] });
    const imagePatchPath = path.join(place, "office-image.patch.yml");
    const pluginPath = path.join(harnessDirectory, "office-image.mjs").replaceAll("\\", "/");
    await fs.writeFile(imagePatchPath,
      `- insert:\n    - id: office-image\n      name: ${JSON.stringify(pluginPath)}\n`, "utf8");
    const env = buildRuntimeEnv({ baseEnv: process.env, home: place, gatewayUrl: gateway.url,
      gatewayKey: gateway.key, model: MODEL, system: "只在需要图片时调用生图工具。",
      imageBrokerUrl: broker.url, imageBrokerKey: broker.key, dshBin });
    const text = await runDsh({ executable: process.execPath, dshBin, patchPath,
      extraPatchPaths: [imagePatchPath], cwd: place, env, model: MODEL,
      prompt: "请生成横版封面图。" });
    assert.equal(text, "图片已生成，作为本轮回复的配图。 ".trim());
    assert.equal(brokerCalls.length, 1);
    assert.deepEqual({ prompt: brokerCalls[0].prompt, size: brokerCalls[0].size, quality: brokerCalls[0].quality },
      { prompt: "蓝色山峦的简洁封面", size: "1536x1024", quality: "high" });
    assert.equal(captured.length, 2);
    assert.equal(captured[0].tools.some((tool) => tool.function?.name === "office_image_generate"), true);
    const toolResults = captured[1].messages.filter((message) => message.role === "tool");
    assert.equal(toolResults.length, 1);
    assert.match(toolResults[0].content, /已生成第 1 张图片/);
    assert.equal(toolResults[0].content.includes("private-generated-image-base64"), false);
    assert.equal(JSON.stringify(captured[1].messages).includes("private-generated-image-base64"), false);
  } finally {
    if (gateway) await gateway.close();
    await broker.close();
    await upstream.close();
    await fs.rm(place, { recursive: true, force: true });
  }
});
