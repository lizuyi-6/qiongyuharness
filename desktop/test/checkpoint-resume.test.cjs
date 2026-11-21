"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { test } = require("node:test");
const { createCheckpointStore, summarizeCheckpoint } = require("../lib/checkpoint.cjs");
const { buildResumeContext, runPipeline } = require("../lib/pipeline.cjs");
const { createDesktopAgent } = require("../agent.cjs");

const ROLES = {
  understand: "glm-5.3",
  execute: "step-5-preview",
  accept: "glm-5.3",
  review: "glm-5.3",
  light: "glm-5.3"
};
const PROMPTS = {
  understand: "你是需求理解专家",
  accept: "你是独立验收专家",
  review: "你是高级复核专家",
  revise: "请按验收意见返工修订"
};
const PIPELINE_CONFIG = {
  enabled: true,
  version: "0123456789ab",
  roles: ROLES,
  prompts: PROMPTS,
  limits: { max_exec_rounds: 2, deadline_ms: 270_000 }
};

test("checkpoint store saves, isolates by userId, summarizes without base64, and clears", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "qo-cp-test-"));
  try {
    const store = createCheckpointStore({ baseDir: dir });
    const origin = "https://office.aethernettech.dev";
    const convId = "conv-123";
    await store.save(origin, convId, {
      runId: "run-1",
      userId: "user-a",
      message: "帮我做一个介绍 MiniMax 的 PPT",
      mode: "pptx",
      model: "qo:pipeline",
      stage: "execute_round_1",
      error: "任务执行阶段超时",
      brief: { tier: "standard", goal: "制作介绍 MiniMax 的 PPT", constraints: [], acceptance: [] },
      checkedDeck: { title: "MiniMax", slides: [{ title: "封面", elements: [] }, { title: "正文", elements: [] }] },
      searchState: {
        searchCount: 2,
        fetchCount: 1,
        queriesHistory: ["MiniMax 2026"],
        evidenceByUrl: [{ url: "https://minimax.io", title: "MiniMax 官网", published_at: "2026-11-01", excerpt: "AGI" }],
        fetchedPages: [["https://minimax.io", "官网正文"]]
      },
      images: [
        { prompt: "科技封面", size: "1536x864", quality: "medium", model: "gpt-image-2.6-flare", width: 1536, height: 864, png_base64: "SECRET_BASE64_BLOB" }
      ]
    });

    const summary = await store.summary(origin, convId, { userId: "user-a" });
    assert.ok(summary);
    assert.equal(summary.stage, "execute_round_1");
    assert.equal(summary.stageLabel, "方案执行");
    assert.equal(summary.imageCount, 1);
    assert.equal(summary.searchCount, 2);
    assert.equal(summary.hasBrief, true);
    assert.equal(summary.hasDraft, true);
    assert.equal(JSON.stringify(summary).includes("SECRET_BASE64_BLOB"), false);
    assert.ok(summary.preserved.some((t) => t.includes("已理解需求")));
    assert.ok(summary.preserved.some((t) => t.includes("已保留 1 张配图")));
    assert.ok(summary.preserved.some((t) => t.includes("2 页 PPT 草稿")));

    // 不同 userId 不得串号读取
    const otherUser = await store.summary(origin, convId, { userId: "user-b" });
    assert.equal(otherUser, null);

    await store.clear(origin, convId);
    assert.equal(await store.load(origin, convId), null);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("runPipeline resumes from execute_round_1 checkpoint, skips understand, and injects draft + images + thinking", async () => {
  const calls = [];
  const checkpoints = [];
  const resumeState = {
    stage: "execute_round_1",
    stageLabel: "方案执行",
    error: "任务执行阶段超时",
    brief: { tier: "standard", goal: "介绍 MiniMax", constraints: ["使用官网配色"], acceptance: ["9页PPT"] },
    tier: "standard",
    checkedDeck: { title: "MiniMax", slides: [{ title: "第1页", elements: [] }] },
    thinkingTail: "Fix slide 9: move #22 to y5.55 h0.6, then emit final JSON.",
    stages: [{ role: "understand", model: "glm-5.3", outcome: "ok", usage: { total_tokens: 120 } }]
  };

  const result = await runPipeline({
    job: { system: "你是资深办公助理", prompt: "配色跟minimax的相差比较大" },
    config: PIPELINE_CONFIG,
    availableModels: Object.values(ROLES),
    resumeState,
    getGeneratedImageCount: () => 2,
    getGeneratedImages: () => [
      { prompt: "MiniMax 珊瑚红科技封面", size: "1536x864" },
      { prompt: "多模态抽象几何插画", size: "1024x1024" }
    ],
    requireSearch: true,
    getSearchCount: () => 2,
    getSearchEvidence: () => "[已检索] MiniMax 官网主色 #FF3763",
    onCheckpoint: (cp) => checkpoints.push(cp),
    runStage: async (stage) => {
      calls.push(stage);
      if (stage.role === "execute") {
        assert.match(stage.system, /断点续跑上下文/);
        assert.match(stage.system, /image_index = 1/);
        assert.match(stage.system, /MiniMax 官网主色 #FF3763/);
        assert.match(stage.system, /Fix slide 9/);
        assert.match(stage.system, /已通过排版校验的 PPT 自由画布草稿/);
        return {
          text: JSON.stringify({ action: "reply", reply: "已按 MiniMax 品牌色完成修改。" }),
          usage: { total_tokens: 500 }
        };
      }
      if (stage.role === "accept") {
        return {
          text: JSON.stringify({ verdict: "pass", issues: [], summary: "通过" }),
          usage: { total_tokens: 100 }
        };
      }
      throw new Error(`unexpected role ${stage.role}`);
    }
  });

  assert.deepEqual(calls.map((c) => c.role), ["execute", "accept"]);
  assert.equal(result.pipeline.verdict, "pass");
  assert.equal(result.pipeline.stages.length, 3);
  assert.ok(checkpoints.some((c) => c.stage === "accept_round_1"));
});

test("runPipeline resumes directly at execute_round_2 without re-running round 1 or accept 1", async () => {
  const calls = [];
  const resumeState = {
    stage: "execute_round_2",
    brief: { tier: "standard", goal: "输出方案", constraints: [], acceptance: [] },
    tier: "standard",
    finalText: JSON.stringify({ action: "reply", reply: "第1轮初稿" }),
    verdict: "fail",
    rounds: 1,
    issues: ["缺少 2026 年最新数据"],
    stages: [
      { role: "understand", model: "glm-5.3", outcome: "ok", usage: {} },
      { role: "execute", model: "step-5-preview", outcome: "ok", usage: {} },
      { role: "accept", model: "glm-5.3", outcome: "fail", usage: {} }
    ]
  };

  const result = await runPipeline({
    job: { system: "系统提示", prompt: "请完成汇报" },
    config: PIPELINE_CONFIG,
    availableModels: Object.values(ROLES),
    resumeState,
    runStage: async (stage) => {
      calls.push(stage.role);
      if (stage.role === "execute") {
        assert.match(stage.system, /缺少 2026 年最新数据/);
        return { text: JSON.stringify({ action: "reply", reply: "第2轮修订稿（含2026年数据）" }), usage: {} };
      }
      if (stage.role === "accept") {
        return { text: JSON.stringify({ verdict: "pass", issues: [], summary: "修订通过" }), usage: {} };
      }
      throw new Error(`unexpected stage ${stage.role}`);
    }
  });

  assert.deepEqual(calls, ["execute", "accept"]);
  assert.equal(result.pipeline.rounds, 2);
  assert.equal(result.pipeline.verdict, "pass");
  assert.match(result.text, /第2轮修订稿/);
});

test("createDesktopAgent resumes complete checkpoint with zero model calls and persists failed mid-run checkpoint", async () => {
  const userData = await fs.mkdtemp(path.join(os.tmpdir(), "qo-agent-cp-"));
  const origin = "https://office.aethernettech.dev";
  const convId = "conv-agent-1";
  let dshRuns = 0;

  try {
    const fakeApp = {
      isPackaged: false,
      getPath: () => userData
    };
    const agent = createDesktopAgent({
      app: fakeApp,
      getWindow: () => null,
      getOrigin: () => origin,
      dependencies: {
        freshModels: async () => ["step-5-preview", "gpt-image-2.6-flare"],
        getCredential: async () => ({ credential: { baseUrl: "https://api.example.com/v1", apiKey: "k-test" } }),
        fetchSessionJson: async () => ({ configured: false }),
        createLoopbackGateway: async () => ({
          url: "http://127.0.0.1:1",
          key: "gw-key",
          usage: { total_tokens: 321 },
          max_tokens_clamped: [],
          finish_reasons: ["stop"],
          truncated: false,
          close: async () => {}
        }),
        nodeExecutable: () => ({ executable: process.execPath, extraEnv: {} }),
        runDsh: async ({ onCheckedDeck }) => {
          dshRuns += 1;
          if (dshRuns === 1) {
            onCheckedDeck?.({ title: "半成品PPT", slides: [{ title: "封面", elements: [] }] });
            throw new Error("桌面 Harness 调用超时，请重试");
          }
          return JSON.stringify({ action: "reply", reply: "断点续跑成功完成！" });
        }
      }
    });

    // 1. 首次运行失败，自动保存包含 checkedDeck 的断点
    await assert.rejects(
      agent.runAgent({
        conversation_id: convId,
        run_id: "run-fail-1",
        user_id: "u-1",
        system: "sys",
        prompt: "做一份 PPT",
        model: "step-5-preview",
        checkpoint_meta: { message: "做一份 PPT", mode: "pptx", material_ids: [] }
      }),
      /超时/
    );

    const cpAfterFail = await agent.getCheckpoint(convId, "u-1");
    assert.ok(cpAfterFail);
    assert.equal(cpAfterFail.hasDraft, true);
    assert.equal(cpAfterFail.hasCompletedResult, false);
    assert.match(cpAfterFail.error, /超时/);

    // 2. 断点续跑成功并写入 stage: "complete" 缓存
    const resumed = await agent.runAgent({
      conversation_id: convId,
      run_id: "run-resume-2",
      user_id: "u-1",
      system: "sys",
      prompt: "做一份 PPT",
      model: "step-5-preview",
      resume: true
    });
    assert.equal(dshRuns, 2);
    assert.match(resumed.text, /断点续跑成功完成/);

    const cpAfterComplete = await agent.getCheckpoint(convId, "u-1");
    assert.ok(cpAfterComplete);
    assert.equal(cpAfterComplete.stage, "complete");
    assert.equal(cpAfterComplete.hasCompletedResult, true);

    // 3. 模拟前端 /complete 网络失败后再次点续跑 -> 0 次模型调用直接返回完整结果
    const instantResumed = await agent.runAgent({
      conversation_id: convId,
      run_id: "run-resume-3",
      user_id: "u-1",
      system: "sys",
      prompt: "做一份 PPT",
      model: "step-5-preview",
      resume: true
    });
    assert.equal(dshRuns, 2);
    assert.equal(instantResumed.resumedFromComplete, true);
    assert.match(instantResumed.text, /断点续跑成功完成/);

    // 4. 保存成功后清理断点
    await agent.clearCheckpoint(convId);
    assert.equal(await agent.getCheckpoint(convId, "u-1"), null);
  } finally {
    await fs.rm(userData, { recursive: true, force: true });
  }
});

test("buildResumeContext formats images, draft, evidence and thinking cleanly", () => {
  const ctx = buildResumeContext({
    resumeState: { stageLabel: "方案执行", error: "网络中断", thinkingTail: "已定好配色 #FF3763" },
    searchEvidence: "2026年最新资料",
    images: [{ prompt: "封面图", size: "1536x864" }]
  });
  assert.match(ctx, /网络中断/);
  assert.match(ctx, /2026年最新资料/);
  assert.match(ctx, /image_index = 1/);
  assert.match(ctx, /已定好配色 #FF3763/);
  assert.ok(summarizeCheckpoint({ stage: "complete", completedResult: { text: "{}" } }).hasCompletedResult);
});

test("invalidateResult keeps billed images and search while dropping the result payload", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "qo-cp-inv-"));
  try {
    const store = createCheckpointStore({ baseDir: dir });
    const origin = "https://office.aethernettech.dev";
    const convId = "conv-inv";
    await store.save(origin, convId, {
      runId: "run-inv",
      userId: "user-a",
      stage: "complete",
      completedResult: { text: "{\"reply\":\"被判死的结果\"}" },
      finalText: "{\"reply\":\"被判死的结果\"}",
      brief: { tier: "standard", goal: "介绍" },
      checkedDeck: { title: "草稿", slides: [{ title: "封面", elements: [] }] },
      thinkingTail: "第 9 页挪一下",
      stages: [{ role: "understand", model: "m", outcome: "ok", usage: {} }],
      issues: ["旧问题"],
      steps: ["读取资料"],
      images: [{ prompt: "封面", png_base64: "BILLED_IMAGE_BLOB" }],
      searchState: { searchCount: 3, evidenceByUrl: [{ url: "https://e.example", title: "权威", excerpt: "x" }] }
    });
    assert.equal(await store.invalidateResult(origin, convId), true);
    const data = await store.load(origin, convId);
    assert.equal(data.completedResult, undefined);
    assert.equal(data.finalText, undefined);
    assert.equal(data.brief, undefined);
    assert.equal(data.checkedDeck, undefined);
    assert.equal(data.thinkingTail, undefined);
    assert.equal(data.stage, "");
    assert.deepEqual(data.stages, []);
    // 已计费的配图与联网搜索记录保留。
    assert.equal(data.images.length, 1);
    assert.equal(data.searchState.searchCount, 3);
    assert.deepEqual(data.steps, ["读取资料"]);
    const summary = await store.summary(origin, convId, { userId: "user-a" });
    assert.equal(summary.hasCompletedResult, false);
    assert.equal(summary.imageCount, 1);
    assert.equal(summary.searchCount, 3);
    assert.equal(await store.invalidateResult(origin, "conv-none"), false);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("invalidateCheckpoint forces regeneration instead of replaying the rejected result", async () => {
  const userData = await fs.mkdtemp(path.join(os.tmpdir(), "qo-agent-inv-"));
  const origin = "https://office.aethernettech.dev";
  const convId = "conv-agent-inv";
  let dshRuns = 0;
  try {
    const agent = createDesktopAgent({
      app: { isPackaged: false, getPath: () => userData },
      getWindow: () => null,
      getOrigin: () => origin,
      dependencies: {
        freshModels: async () => ["step-5-preview", "gpt-image-2.6-flare"],
        getCredential: async () => ({ credential: { baseUrl: "https://api.example.com/v1", apiKey: "k-test" } }),
        fetchSessionJson: async () => ({ configured: false }),
        createLoopbackGateway: async () => ({
          url: "http://127.0.0.1:1", key: "gw-key", usage: { total_tokens: 7 },
          max_tokens_clamped: [], finish_reasons: ["stop"], truncated: false, close: async () => {}
        }),
        nodeExecutable: () => ({ executable: process.execPath, extraEnv: {} }),
        runDsh: async () => {
          dshRuns += 1;
          return JSON.stringify({ action: "reply", reply: "全新生成的结果" });
        }
      }
    });
    const job = (runId, extra = {}) => ({
      conversation_id: convId, run_id: runId, user_id: "u-1",
      system: "sys", prompt: "做一份 PPT", model: "step-5-preview",
      checkpoint_meta: { message: "做一份 PPT", mode: "pptx", material_ids: [] }, ...extra
    });
    const first = await agent.runAgent(job("run-inv-1"));
    assert.equal(dshRuns, 1);
    const cp = await agent.getCheckpoint(convId, "u-1");
    assert.equal(cp.stage, "complete");
    assert.equal(cp.hasCompletedResult, true);

    // 保存被拒前：断点续跑 0 次模型调用，原样重放。
    const replay = await agent.runAgent(job("run-inv-2", { resume: true }));
    assert.equal(dshRuns, 1);
    assert.equal(replay.resumedFromComplete, true);

    // 前端收到 /complete 4xx 后作废结果载荷。
    assert.equal(await agent.invalidateCheckpoint(convId), true);
    const invalidated = await agent.getCheckpoint(convId, "u-1");
    assert.equal(invalidated.hasCompletedResult, false);

    // 再点续跑：重新生成而不是重放同一份被判死的结果。
    const regenerated = await agent.runAgent(job("run-inv-3", { resume: true }));
    assert.equal(dshRuns, 2);
    assert.equal(regenerated.resumedFromComplete, undefined);
    assert.match(regenerated.text, /全新生成的结果/);

    assert.equal(await agent.invalidateCheckpoint("conv-none"), false);
  } finally {
    await fs.rm(userData, { recursive: true, force: true });
  }
});
