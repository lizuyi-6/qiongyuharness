const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const Module = require("node:module");
const { PASSTHROUGH_ENV } = require("../lib/runtime-env.cjs");
const { DEADLINE_MS, STAGE_TIMEOUT_MS } = require("../lib/pipeline.cjs");

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
  assert.equal(stage.timeoutMs, STAGE_TIMEOUT_MS);
  assert.match(stage.system, /生图工具使用说明：/);
  assert.deepEqual(result.images, [image]);
  assert.notEqual(result.images, captures.brokerArgs[0].images);
  assert.equal(captures.brokerArgs[0].model, "gpt-image-2.6-flare");
  assert.equal(captures.brokerArgs[0].deadlineAt, 10_000 + STAGE_TIMEOUT_MS);
  await captures.brokerArgs[0].generate({ model: "gpt-image-2.6-flare", prompt: "x", size: "1024x1024", quality: "low" },
    { timeoutMs: 120_000 });
  generation = true;
  assert.equal(generation, true);
  assert.equal(captures.brokerClosed, true);
});

test("生图未启用时模型列表、system、allowedTools、env 和阶段 timeout 不变", async (t) => {
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
    assert.equal(stage.timeoutMs, STAGE_TIMEOUT_MS);
    assert.equal(stage.brokerUrl, undefined);
    assert.equal(stage.brokerKey, undefined);
    assert.equal(stage.extraPatchPaths, undefined);
    const baseKeys = ["DSH_AGENTS_HOME", "DSH_HOME", "DSH_SYSTEM_PROMPT", "HOME", "NO_PROXY",
      "QO_DSH_EXECUTABLE", "QO_DSH_PLUGIN_CONTEXT", "QO_RUNTIME_KEY", "QO_RUNTIME_MODEL",
      "QO_RUNTIME_URL", "QO_SERVER_DATE", "QO_SERVER_TIMESTAMP", "QO_SERVER_TIME_DISPLAY",
      "QO_SERVER_YEAR", "USERPROFILE"];
    assert.deepEqual(stage.envKeys, [...baseKeys, ...PASSTHROUGH_ENV.filter((name) => typeof process.env[name] === "string")].sort());
  });
});

test("qo:pipeline 按 image_tool 为执行/轻量阶段开放图片工具", async () => {
  const responses = [
    JSON.stringify({ tier: "light", goal: "answer", constraints: [], acceptance: ["reply exists"] }),
    JSON.stringify({ action: "reply", reply: "result" })
  ];
  const available = ["chat-model", ...Object.values(roles), ...imageModels];
  const base = agentHarness({ models: available, responses: [...responses] });
  const withFlag = agentHarness({ models: available, responses: [...responses] });
  await base.agent.runAgent({ system: "same pipeline system", prompt: "request", model: "qo:pipeline", image_tool: false });
  await withFlag.agent.runAgent({ system: "same pipeline system", prompt: "request", model: "qo:pipeline", image_tool: true });
  // 0.2.6 起智能分工支持按需生图：开启时创建一次性 broker，全程 deadline 作废线。
  assert.equal(base.captures.brokerCreates, 0);
  assert.equal(withFlag.captures.brokerCreates, 1);
  assert.equal(withFlag.captures.brokerArgs[0].deadlineAt, 10_000 + DEADLINE_MS);
  const [baseUnderstand, baseLight] = base.captures.stages;
  const [flagUnderstand, flagLight] = withFlag.captures.stages;
  // 理解阶段使用阶段提示词，不接触 broker，工具列表不变。
  assert.equal(flagUnderstand.system, "understand");
  assert.equal(flagUnderstand.system, baseUnderstand.system);
  assert.equal(flagUnderstand.brokerUrl, undefined);
  assert.deepEqual(flagUnderstand.allowedTools, baseUnderstand.allowedTools);
  // 轻量执行阶段在 job.system 上追加生图说明，注入 broker 环境并开放 office_image_generate。
  assert.equal(baseLight.system, "same pipeline system");
  assert.match(flagLight.system, /^same pipeline system\n\n生图工具使用说明：/);
  assert.equal(baseLight.brokerUrl, undefined);
  assert.deepEqual(baseLight.allowedTools, baseUnderstand.allowedTools);
  assert.equal(flagLight.brokerUrl, "http://127.0.0.1:22222");
  assert.equal(flagLight.brokerKey, "one-time-broker-key");
  assert.deepEqual(flagLight.allowedTools, [...baseLight.allowedTools, "office_image_generate"]);
});

test("DSH 失败时错误原样上抛且 broker 在 finally 关闭", async () => {
  const { agent, captures } = agentHarness({ brokerImages: [image], fail: "dsh 失败" });
  // 0.2.17 起已生成的图片由断点机制持久化，失败不再追加计费提示，错误原样上抛。
  await assert.rejects(agent.runAgent({ system: "s", prompt: "p", model: "chat-model", image_tool: true }),
    (error) => error.message === "dsh 失败");
  assert.equal(captures.brokerClosed, true);
});
