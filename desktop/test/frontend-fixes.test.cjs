const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const Module = require("node:module");

const { runPipeline } = require("../lib/pipeline.cjs");

const REPO_ROOT = path.resolve(__dirname, "..", "..");

function loadAgentWithMockElectron() {
  const originalLoad = Module._load;
  Module._load = function mockElectron(request, parent, isMain) {
    if (request === "electron") {
      return {
        net: { fetch: async () => { throw new Error("unexpected network call"); } },
        safeStorage: { isEncryptionAvailable: () => true }
      };
    }
    return originalLoad.call(this, request, parent, isMain);
  };
  try {
    const resolved = require.resolve("../agent.cjs");
    delete require.cache[resolved];
    return require("../agent.cjs");
  } finally {
    Module._load = originalLoad;
  }
}

test("Markdown 渲染支持常用办公语法并严格防 XSS", async () => {
  const mdUrl = pathToFileURL(path.join(REPO_ROOT, "server", "static", "markdown.js")).href;
  const { renderMarkdown } = await import(mdUrl);

  const sample = [
    "# 汇报标题",
    "",
    "正文包含 **加粗**、*斜体*、~~删除~~ 与 `行内代码`。",
    "",
    "- 条目一",
    "- 条目二",
    "",
    "1. 第一步",
    "2. 第二步",
    "",
    "> 重要提醒",
    "",
    "| 项目 | 状态 |",
    "| --- | --- |",
    "| 接口修复 | 完成 |",
    "",
    "```js",
    "const x = '<script>alert(1)</script>';",
    "```",
    "",
    "安全链接：[官网](https://office.aethernettech.dev)",
    "恶意链接：[点击](javascript:alert(1))",
    "原始标签：<img src=x onerror=alert(1)><script>evil()</script>"
  ].join("\n");

  const html = renderMarkdown(sample);
  assert.match(html, /<h3 class="md-heading">汇报标题<\/h3>/);
  assert.match(html, /<strong>加粗<\/strong>/);
  assert.match(html, /<em>斜体<\/em>/);
  assert.match(html, /<del>删除<\/del>/);
  assert.match(html, /<code>行内代码<\/code>/);
  assert.match(html, /<ul><li>条目一<\/li><li>条目二<\/li><\/ul>/);
  assert.match(html, /<ol><li>第一步<\/li><li>第二步<\/li><\/ol>/);
  assert.match(html, /<blockquote>重要提醒<\/blockquote>/);
  assert.match(html, /<table class="md-table">.*<th>项目<\/th>.*<td>接口修复<\/td>/);
  assert.match(html, /<pre class="md-code" data-lang="js"><code>const x = &#39;&lt;script&gt;alert\(1\)&lt;\/script&gt;&#39;;<\/code><\/pre>/);
  assert.match(html, /<a href="https:\/\/office\.aethernettech\.dev" target="_blank" rel="noopener noreferrer">官网<\/a>/);
  assert.doesNotMatch(html, /href=["']?javascript:/i);
  assert.match(html, /\[点击\]\(javascript:alert\(1\)\)/);
  assert.doesNotMatch(html, /<script>/i);
  assert.doesNotMatch(html, /<img\b/i);
  assert.match(html, /&lt;img src=x onerror=alert\(1\)&gt;/);
});

test("智能分工流水线响应 AbortSignal 立即中止且不执行后续阶段", async () => {
  const controller = new AbortController();
  const seenRoles = [];
  const config = {
    enabled: true,
    version: "0123456789ab",
    roles: {
      understand: "glm-5.3",
      execute: "step-5-preview",
      accept: "glm-5.3",
      review: "k3.1-pro",
      light: "deepseek-v4.1-flash-1017"
    },
    limits: { max_exec_rounds: 2, deadline_ms: 270_000 },
    prompts: { understand: "u", accept: "a", review: "r", revise: "v" }
  };

  await assert.rejects(() => runPipeline({
    job: { system: "sys", prompt: "do work" },
    config,
    availableModels: Object.values(config.roles),
    signal: controller.signal,
    runStage: async ({ role, signal }) => {
      seenRoles.push(role);
      controller.abort();
      if (signal.aborted) {
        throw Object.assign(new Error("已停止生成"), { code: "ABORTED" });
      }
      return { text: "{}", usage: {} };
    }
  }), /已停止生成/);

  assert.deepEqual(seenRoles, ["understand"]);
});

test("桌面 agent.cancelAgent() 能中止运行中的任务", async () => {
  const { createDesktopAgent } = loadAgentWithMockElectron();
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), "qo-cancel-test-"));
  let releaseStage;
  const stageStarted = new Promise((resolve) => { releaseStage = resolve; });
  let observedAbort = false;

  const agent = createDesktopAgent({
    app: { getPath: () => userData, isPackaged: false },
    getWindow: () => ({ isDestroyed: () => false }),
    getOrigin: () => "https://office.aethernettech.dev",
    dependencies: {
      freshModels: async () => ["glm-5.3"],
      getCredential: async () => ({
        user: { id: "u1" },
        credential: { baseUrl: "https://aethernettech.dev/v1", apiKey: "sk-test" }
      }),
      createLoopbackGateway: async () => ({
        url: "http://127.0.0.1:1/v1",
        key: "tmp",
        usage: {},
        max_tokens_clamped: [],
        finish_reasons: [],
        truncated: false,
        quota_error: null,
        close: async () => {}
      }),
      nodeExecutable: () => ({ executable: process.execPath, extraEnv: {} }),
      runDsh: ({ signal }) => new Promise((_resolve, reject) => {
        releaseStage();
        signal.addEventListener("abort", () => {
          observedAbort = true;
          reject(Object.assign(new Error("已停止生成"), { code: "ABORTED" }));
        }, { once: true });
      })
    }
  });

  try {
    const pending = agent.runAgent({ system: "sys", prompt: "hi", model: "glm-5.3" });
    await stageStarted;
    assert.equal(agent.cancelAgent(), true);
    await assert.rejects(() => pending, /已停止生成/);
    assert.equal(observedAbort, true);
    assert.equal(agent.cancelAgent(), false);
  } finally {
    fs.rmSync(userData, { recursive: true, force: true });
  }
});

test("前端文稿页桌面回传路径与后端 files.py 路由一致", () => {
  const appJs = fs.readFileSync(path.join(REPO_ROOT, "server", "static", "app.js"), "utf8");
  const filesPy = fs.readFileSync(path.join(REPO_ROOT, "server", "files.py"), "utf8");

  assert.match(filesPy, /@router\.post\("\/agent\/plan\/complete"\)/);
  assert.match(filesPy, /@router\.post\("\/assignments\/\{key\}\/sessions\/complete"\)/);
  assert.match(appJs, /return post\(base\+"\/complete",\{/);
  assert.doesNotMatch(appJs, /post\(base\+"\/"\+encodeURIComponent\(prepared\.run_id\)\+"\/complete"/);
});
