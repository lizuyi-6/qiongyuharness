const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const Module = require("node:module");
const { PASSTHROUGH_ENV } = require("../lib/runtime-env.cjs");

const roles = { understand: "understand-model", execute: "execute-model", accept: "accept-model",
  review: "review-model", light: "light-model" };
const imageModels = ["gpt-image-2.6-sunburst", "gpt-image-2.6-flare"];
const pipelineConfig = {
  enabled: true, version: "0123456789ab", roles,
  prompts: { understand: "understand", accept: "accept", review: "review", revise: "revise" },
  limits: { max_exec_rounds: 2, deadline_ms: 270_000 }
};
const usage = { prompt_tokens: 1, completion_tokens: 2, total_tokens: 3,
  max_tokens_clamped: [], finish_reasons: ["stop"], truncated: false };
const image = { png_base64: "only-in-employee-process", width: 1024, height: 1024,
  model: "gpt-image-2.6-flare", size: "1024x1024", quality: "medium", prompt: "test", usage: {} };

function loadAgent() {
  const originalLoad = Module._load;
  Module._load = function (request, parent, isMain) {
    if (request === "electron") return { net: { fetch: async () => { throw new Error("unexpected network"); } }, safeStorage: {} };
    return originalLoad.call(this, request, parent, isMain);
  };
  try { return require("../agent.cjs").createDesktopAgent; }
  finally { Module._load = originalLoad; }
}

function agentHarness({ models = ["chat-model", ...imageModels], responses = ["reply"], fail = null,
  brokerImages = [], onGenerate = () => {} } = {}) {
  const captures = { stages: [], gateways: [], brokerArgs: [], brokerClosed: false, brokerCreates: 0 };
  const createDesktopAgent = loadAgent();
  const agent = createDesktopAgent({
    app: { getPath: () => require("node:os").tmpdir(), isPackaged: false },
    getWindow: () => null,
    getOrigin: () => "https://qo.example",
    emit: () => {},
    dependencies: {
      now: () => 10_000,
      freshModels: async () => models,
      getCredential: async () => ({ credential: { baseUrl: "https://models.example/v1", apiKey: "employee-key" } }),
      fetchSessionJson: async () => pipelineConfig,
      createLoopbackGateway: async (options) => {
        captures.gateways.push(options);
        return { url: "http://127.0.0.1:11111/v1", key: "gateway-only-key", usage,
          max_tokens_clamped: [], finish_reasons: ["stop"], truncated: false, quota_error: null,
          close: async () => {} };
      },
      runDsh: async (options) => {
        const patchContent = options.extraPatchPaths?.length ? await fs.readFile(options.extraPatchPaths[0], "utf8") : "";
        captures.stages.push({
          system: options.env.DSH_SYSTEM_PROMPT,
          timeoutMs: options.timeoutMs,
          envKeys: Object.keys(options.env).sort(),
          brokerUrl: options.env.QO_IMAGE_BROKER_URL,
          brokerKey: options.env.QO_IMAGE_BROKER_KEY,
          dshExecutable: options.env.QO_DSH_EXECUTABLE,
          allowedTools: captures.gateways.at(-1).allowedTools,
          contextPatchPath: options.contextPatchPath,
          extraPatchPaths: options.extraPatchPaths,
          patchContent
        });
        if (fail) throw new Error(fail);
        return responses.shift() ?? "reply";
      },
      createImageBroker: async (options) => {
        captures.brokerCreates += 1;
        captures.brokerArgs.push(options);
        return { url: "http://127.0.0.1:22222", key: "one-time-broker-key", images: brokerImages,
          close: async () => { captures.brokerClosed = true; } };
      },
      imageGenerate: async (input, options) => { onGenerate(input, options); return image; },
      nodeExecutable: () => ({ executable: process.execPath, extraEnv: {} })
    }
  });
  return { agent, captures };
}

test("models() 从聊天 models 隔离生图模型并按字母序放入 image_models", async () => {
  const { agent } = agentHarness();
  const value = await agent.models();
  assert.deepEqual(value.models, ["chat-model"]);
  assert.deepEqual(value.image_models, ["gpt-image-2.6-flare", "gpt-image-2.6-sunburst"]);
});

test("image_tool 启用时精确开放办公工具、注入 broker patch/env 并关闭 broker", async () => {
  const { agent, captures } = agentHarness({ brokerImages: [image] });
  let generation;
  const result = await agent.runAgent({ system: "base system", prompt: "请求插图", model: "chat-model",
    context: { documents: [] }, image_tool: true });
  const stage = captures.stages[0];
  assert.deepEqual(captures.gateways[0].allowedTools, ["office_context_read", "office_image_generate"]);
  assert.equal(stage.brokerUrl, "http://127.0.0.1:22222");
  assert.equal(stage.brokerKey, "one-time-broker-key");
  assert.equal(stage.dshExecutable.endsWith("@deepseek-ai\\dsh\\lib\\bin.js") ||
    stage.dshExecutable.endsWith("@deepseek-ai/dsh/lib/bin.js"), true);
  assert.equal(stage.extraPatchPaths.length, 1);
  assert.match(stage.patchContent, /office-image\.mjs/);
  assert.equal(stage.timeoutMs, 240_000);
  assert.match(stage.system, /生图工具使用说明：/);
  assert.deepEqual(result.images, [image]);
  assert.notEqual(result.images, captures.brokerArgs[0].images);
  assert.equal(captures.brokerArgs[0].model, "gpt-image-2.6-flare");
  assert.equal(captures.brokerArgs[0].deadlineAt, 250_000);
  await captures.brokerArgs[0].generate({ model: "gpt-image-2.6-flare", prompt: "x", size: "1024x1024", quality: "low" },
    { timeoutMs: 120_000 });
  generation = true;
  assert.equal(generation, true);
  assert.equal(captures.brokerClosed, true);
});

test("生图未启用时模型列表、system、allowedTools、env 和 95 秒 timeout 不变", async (t) => {
  const context = { documents: [] };
  for (const scenario of [
    { name: "image_tool 缺失", job: {} },
    { name: "image_tool 为 false", job: { image_tool: false } },
    { name: "没有生图模型", job: { image_tool: true }, models: ["chat-model"] }
  ]) await t.test(scenario.name, async () => {
    const { agent, captures } = agentHarness({ models: scenario.models || ["chat-model", ...imageModels] });
    await agent.runAgent({ system: "same system", prompt: "request", model: "chat-model", context, ...scenario.job });
    const stage = captures.stages[0];
    assert.equal(captures.brokerCreates, 0);
    assert.deepEqual(captures.gateways[0].allowedTools, ["office_context_read"]);
    assert.equal(stage.system, "same system");
    assert.equal(stage.timeoutMs, 95_000);
    assert.equal(stage.brokerUrl, undefined);
    assert.equal(stage.brokerKey, undefined);
    assert.equal(stage.extraPatchPaths, undefined);
    const baseKeys = ["DSH_AGENTS_HOME", "DSH_HOME", "DSH_SYSTEM_PROMPT", "HOME", "NO_PROXY",
      "QO_DSH_EXECUTABLE", "QO_DSH_PLUGIN_CONTEXT", "QO_RUNTIME_KEY", "QO_RUNTIME_MODEL",
      "QO_RUNTIME_URL", "USERPROFILE"];
    assert.deepEqual(stage.envKeys, [...baseKeys, ...PASSTHROUGH_ENV.filter((name) => typeof process.env[name] === "string")].sort());
  });
});

test("qo:pipeline 不开放图片工具且与原 system 和阶段 timeout 相同", async () => {
  const responses = [
    JSON.stringify({ tier: "light", goal: "answer", constraints: [], acceptance: ["reply exists"] }),
    JSON.stringify({ action: "reply", reply: "result" })
  ];
  const available = ["chat-model", ...Object.values(roles), ...imageModels];
  const base = agentHarness({ models: available, responses: [...responses] });
  const withFlag = agentHarness({ models: available, responses: [...responses] });
  await base.agent.runAgent({ system: "same pipeline system", prompt: "request", model: "qo:pipeline", image_tool: false });
  await withFlag.agent.runAgent({ system: "same pipeline system", prompt: "request", model: "qo:pipeline", image_tool: true });
  assert.equal(withFlag.captures.brokerCreates, 0);
  assert.deepEqual(withFlag.captures.stages.map((stage) => [stage.system, stage.timeoutMs, stage.envKeys,
    stage.allowedTools, stage.extraPatchPaths]), base.captures.stages.map((stage) => [stage.system, stage.timeoutMs,
    stage.envKeys, stage.allowedTools, stage.extraPatchPaths]));
});

test("DSH 失败时 broker 已生成的图片标明已计费并在 finally 关闭", async () => {
  const { agent, captures } = agentHarness({ brokerImages: [image], fail: "dsh 失败" });
  await assert.rejects(agent.runAgent({ system: "s", prompt: "p", model: "chat-model", image_tool: true }),
    /dsh 失败（本轮已生成 1 张图片并已计费，因执行失败未能保存）$/);
  assert.equal(captures.brokerClosed, true);
});
