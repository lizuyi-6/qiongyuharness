const test = require("node:test");
const assert = require("node:assert/strict");
const { DEADLINE_MS, MIN_REMAINING_MS, STAGE_TIMEOUT_MS, runPipeline } = require("../lib/pipeline.cjs");

const roles = {
  understand: "model-understand",
  execute: "model-execute",
  accept: "model-accept",
  review: "model-review",
  light: "model-light"
};
const prompts = { understand: "understand prompt", accept: "accept prompt", review: "review prompt", revise: "revise prompt" };
const config = { enabled: true, version: "0123456789ab", roles, prompts, limits: { max_exec_rounds: 2, deadline_ms: 270_000 } };
const models = Object.values(roles);
const usage = { prompt_tokens: 1, completion_tokens: 2, total_tokens: 3 };
const brief = (tier = "standard") => JSON.stringify({ tier, goal: "deliver", constraints: [], acceptance: ["reply exists"] });
const deliverable = (reply = "answer") => JSON.stringify({ action: "reply", reply });
const assessment = (verdict, issues = []) => JSON.stringify({ verdict, issues, summary: verdict });

function runner(responses, onCall = () => {}) {
  const calls = [];
  return {
    calls,
    runStage: async (stage) => {
      calls.push(stage);
      onCall(stage, calls.length);
      const response = responses.shift();
      if (response instanceof Error) throw response;
      if (response?.error) throw response.error;
      return { text: response?.text ?? response, usage: response?.usage || usage };
    }
  };
}

function invoke(responses, options = {}) {
  const instance = runner(responses, options.onCall);
  return {
    calls: instance.calls,
    result: runPipeline({
      job: options.job || { system: "system", prompt: "original request", context: null },
      config: options.config || config,
      availableModels: options.availableModels || models,
      runStage: instance.runStage,
      emit: options.emit,
      now: options.now || (() => options.time?.value ?? 0),
      startedAt: options.startedAt ?? 0
    })
  };
}

test("light 任务只执行轻量模型并如实跳过验收", async () => {
  const { calls, result } = invoke([{ text: brief("light") }, { text: deliverable("quick"), usage }], {
    job: { system: "system", prompt: "quick request", context: { files: [] } }
  });
  const output = await result;
  assert.equal(output.text, deliverable("quick"));
  assert.equal(output.pipeline.tier, "light");
  assert.equal(output.pipeline.verdict, "skipped");
  assert.equal(output.pipeline.review, "not_needed");
  assert.deepEqual(output.pipeline.stages.map((stage) => [stage.role, stage.outcome]), [["understand", "ok"], ["light", "ok"]]);
  assert.equal(calls[0].withContext, false);
  assert.equal(calls[1].withContext, true);
});

test("standard 任务一轮执行并通过验收", async () => {
  const { calls, result } = invoke([{ text: brief() }, { text: deliverable() }, { text: assessment("pass") }]);
  const output = await result;
  assert.equal(output.pipeline.verdict, "pass");
  assert.equal(output.pipeline.review, "not_needed");
  assert.equal(output.pipeline.rounds, 1);
  assert.deepEqual(output.pipeline.stages.map((stage) => stage.role), ["understand", "execute", "accept"]);
  assert.equal(output.usage.total_tokens, 9);
  assert.equal(calls[0].timeoutMs, STAGE_TIMEOUT_MS);
});

test("验收 fail 后返工一轮并通过", async () => {
  const { calls, result } = invoke([
    { text: brief() }, { text: deliverable("first") }, { text: assessment("fail", ["缺少标题"]) },
    { text: deliverable("revised") }, { text: assessment("pass") }
  ]);
  const output = await result;
  assert.equal(output.text, deliverable("revised"));
  assert.equal(output.pipeline.verdict, "pass");
  assert.equal(output.pipeline.rounds, 2);
  assert.deepEqual(output.pipeline.stages.map((stage) => stage.outcome), ["ok", "ok", "fail", "ok", "pass"]);
  assert.match(calls[3].system, /缺少标题/);
  assert.match(calls[3].system, /first/);
});

test("两轮验收 fail 后复核 pass 采用复核结论", async () => {
  const { result } = invoke([
    { text: brief() }, { text: deliverable("first") }, { text: assessment("fail", ["缺少标题"]) },
    { text: deliverable("second") }, { text: assessment("fail", ["仍缺标题"]) },
    { text: assessment("pass") }
  ]);
  const output = await result;
  assert.equal(output.pipeline.verdict, "pass");
  assert.equal(output.pipeline.review, "pass");
  assert.equal(output.pipeline.rounds, 2);
  assert.equal(output.pipeline.stages.at(-1).role, "review");
  assert.equal(output.pipeline.stages.at(-1).outcome, "pass");
});

test("两轮验收 fail 且复核 fail 时保留 fail", async () => {
  const { result } = invoke([
    { text: brief() }, { text: deliverable("first") }, { text: assessment("fail", ["缺少标题"]) },
    { text: deliverable("second") }, { text: assessment("fail", ["仍缺标题"]) },
    { text: assessment("fail", ["复核确认缺陷"]) }
  ]);
  const output = await result;
  assert.equal(output.pipeline.verdict, "fail");
  assert.equal(output.pipeline.review, "fail");
  assert.equal(output.pipeline.stages.at(-1).outcome, "fail");
  assert.ok(output.pipeline.issues.includes("复核确认缺陷"));
});

test("critical 任务验收 pass 后仍进入复核", async () => {
  const { calls, result } = invoke([
    { text: brief("critical") }, { text: deliverable() }, { text: assessment("pass") }, { text: assessment("pass") }
  ]);
  const output = await result;
  assert.equal(output.pipeline.verdict, "pass");
  assert.equal(output.pipeline.review, "pass");
  assert.deepEqual(calls.map((call) => call.role), ["understand", "execute", "accept", "review"]);
});

test("critical 任务缺少复核模型时保留验收结论并标记 unavailable", async () => {
  const { result } = invoke([
    { text: brief("critical") }, { text: deliverable() }, { text: assessment("pass") }
  ], { availableModels: models.filter((model) => model !== roles.review) });
  const output = await result;
  assert.equal(output.pipeline.verdict, "pass");
  assert.equal(output.pipeline.review, "unavailable");
  assert.ok(output.pipeline.issues.includes("复核模型未开放，保留验收结果"));
  assert.equal(output.pipeline.stages.some((stage) => stage.role === "review"), false);
});

test("缺少轻量模型时走标准路径并记录回退原因", async () => {
  const { calls, result } = invoke([
    { text: brief("light") }, { text: deliverable() }, { text: assessment("pass") }
  ], { availableModels: models.filter((model) => model !== roles.light) });
  const output = await result;
  assert.equal(output.pipeline.tier, "light");
  assert.equal(output.pipeline.verdict, "pass");
  assert.ok(output.pipeline.issues.includes("轻量模型未开放，已按标准路径执行"));
  assert.deepEqual(calls.map((call) => call.role), ["understand", "execute", "accept"]);
});

test("理解、执行或验收模型缺失时在运行阶段前拒绝", async (t) => {
  for (const role of ["understand", "execute", "accept"]) {
    await t.test(role, async () => {
      const { calls, result } = invoke([], { availableModels: models.filter((model) => model !== roles[role]) });
      await assert.rejects(result, new RegExp(`角色：${role === "understand" ? "理解" : role === "execute" ? "执行" : "验收"}`));
      assert.equal(calls.length, 0);
    });
  }
});

test("理解 JSON 连续两次格式错误时中止", async () => {
  const { calls, result } = invoke(["not json", "still not json"]);
  await assert.rejects(result, /需求理解阶段返回格式无效/);
  assert.equal(calls.length, 2);
  assert.deepEqual(calls.map((call) => call.role), ["understand", "understand"]);
});

test("执行阶段纯文本按兜底交付，验收格式错误重试一次", async () => {
  const { result } = invoke([
    { text: brief() }, "bad execution", { text: deliverable() }, { text: assessment("pass") }
  ]);
  const output = await result;
  assert.equal(output.pipeline.verdict, "pass");
  assert.deepEqual(output.pipeline.stages.map((stage) => [stage.role, stage.outcome]), [
    ["understand", "ok"], ["execute", "ok"], ["accept", "format_error"], ["accept", "pass"]
  ]);
});

test("执行后剩余时间少于 60 秒时不启动验收并按 fail 返回交付物", async () => {
  const time = { value: 0 };
  const { calls, result } = invoke([
    { text: brief() }, { text: deliverable() }
  ], { time, onCall: (_stage, count) => { if (count === 2) time.value = DEADLINE_MS - MIN_REMAINING_MS / 2; } });
  const output = await result;
  assert.equal(output.text, deliverable());
  assert.equal(output.pipeline.verdict, "fail");
  assert.ok(output.pipeline.issues.includes("验收超时未完成"));
  assert.deepEqual(calls.map((call) => call.role), ["understand", "execute"]);
});

test("流水线阶段调用总数封顶 8 次", async () => {
  const { calls, result } = invoke([
    "bad brief", { text: brief() },
    "bad execution", { text: deliverable("first") },
    "bad assessment", { text: assessment("fail", ["缺少标题"]) },
    "bad revision", { text: deliverable("second") }
  ]);
  const output = await result;
  assert.equal(calls.length, 8);
  assert.equal(output.pipeline.stages.length, 8);
  assert.equal(output.pipeline.verdict, "fail");
  assert.ok(output.pipeline.issues.some((issue) => issue.includes("预算已用尽")));
});

test("quota error 在当前阶段立即中止", async () => {
  const quota = Object.assign(new Error("模型额度已用尽"), { quota_error: { type: "rate_limit_exceeded" } });
  const { calls, result } = invoke([{ text: brief() }, quota]);
  await assert.rejects(result, (error) => error === quota);
  assert.equal(calls.length, 2);
  assert.deepEqual(calls.map((call) => call.role), ["understand", "execute"]);
});

function fakeDesktopAgent({ responses, quotaAt = -1 }) {
  const Module = require("node:module");
  const originalLoad = Module._load;
  Module._load = function (request, parent, isMain) {
    if (request === "electron") return { net: { fetch: async () => { throw new Error("unexpected network"); } }, safeStorage: {} };
    return originalLoad.call(this, request, parent, isMain);
  };
  let createDesktopAgent;
  try {
    ({ createDesktopAgent } = require("../agent.cjs"));
  } finally {
    Module._load = originalLoad;
  }

  const gatewayModels = [];
  const dshModels = [];
  const payload = {
    ...config,
    prompts: { ...prompts }
  };
  const agent = createDesktopAgent({
    app: { getPath: () => require("node:os").tmpdir(), isPackaged: false },
    getWindow: () => null,
    getOrigin: () => "https://qo.example",
    dependencies: {
      freshModels: async () => models,
      getCredential: async () => ({ credential: { baseUrl: "https://gateway.example/v1", apiKey: "fake-key" } }),
      fetchSessionJson: async (_origin, url) => {
        assert.equal(url, "https://qo.example/api/v1/me/pipeline");
        return payload;
      },
      createLoopbackGateway: async ({ model }) => {
        gatewayModels.push(model);
        const index = gatewayModels.length - 1;
        return {
          url: "http://127.0.0.1:10000/v1", key: "local-key", usage,
          max_tokens_clamped: [], finish_reasons: ["stop"], truncated: false,
          quota_error: index === quotaAt ? { type: "rate_limit_exceeded", code: "quota_exhausted" } : null,
          close: async () => {}
        };
      },
      runDsh: async ({ model }) => {
        dshModels.push(model);
        return responses.shift();
      },
      nodeExecutable: () => ({ executable: process.execPath, extraEnv: {} }),
      requestUsage: async () => ({ state: "unavailable" })
    }
  });
  return { agent, gatewayModels, dshModels };
}

test("桌面 agent 为每阶段单独建立钉住角色模型的 gateway", async () => {
  const { agent, gatewayModels, dshModels } = fakeDesktopAgent({
    responses: [brief(), deliverable(), assessment("pass")]
  });
  const result = await agent.runAgent({ system: "system", prompt: "request", model: "qo:pipeline" });
  assert.deepEqual(gatewayModels, [roles.understand, roles.execute, roles.accept]);
  assert.deepEqual(dshModels, gatewayModels);
  assert.equal(result.pipeline.verdict, "pass");
  assert.equal(result.model, `智能分工:${roles.execute}`);
});

test("桌面 agent 的普通模型路径仍使用一个钉住模型的 gateway", async () => {
  const selected = roles.execute;
  const { agent, gatewayModels, dshModels } = fakeDesktopAgent({ responses: [deliverable("ordinary")] });
  const result = await agent.runAgent({ system: "system", prompt: "request", model: selected });
  assert.deepEqual(gatewayModels, [selected]);
  assert.deepEqual(dshModels, [selected]);
  assert.equal(result.model, selected);
  assert.equal(result.text, deliverable("ordinary"));
  assert.deepEqual(result.usage.max_tokens_clamped, []);
  assert.deepEqual(result.finish_reasons, ["stop"]);
  assert.equal(result.truncated, false);
  assert.equal(Object.hasOwn(result, "pipeline"), false);
});

test("models() 只公开智能分工模型角色名，不返回阶段提示词", async () => {
  const { agent } = fakeDesktopAgent({ responses: [] });
  const result = await agent.models();
  assert.deepEqual(result.pipeline, { available: true, roles });
  assert.deepEqual(result.models, models);
  assert.equal(Object.hasOwn(result.pipeline, "prompts"), false);
});

test("桌面 agent 在 DSH 返回前后发现 quota_error 时不启动后续阶段", async () => {
  const { agent, gatewayModels, dshModels } = fakeDesktopAgent({ responses: [brief()], quotaAt: 0 });
  await assert.rejects(agent.runAgent({ system: "system", prompt: "request", model: "qo:pipeline" }), /额度已用尽/);
  assert.deepEqual(gatewayModels, [roles.understand]);
  assert.deepEqual(dshModels, [roles.understand]);
});
