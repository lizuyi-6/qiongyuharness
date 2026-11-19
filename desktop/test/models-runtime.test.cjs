const test = require("node:test");
const assert = require("node:assert/strict");
const { listModelsWithRefresh, parseModels, modelServiceError } = require("../lib/models.cjs");
const { buildDshArgs, buildRuntimeEnv, PASSTHROUGH_ENV, resolveRuntimePaths } = require("../lib/runtime-env.cjs");
const path = require("node:path");
const { sessionFetch } = require("../lib/session-cookie.cjs");

test("模型列表解析只返回去重后的模型 id", () => {
  assert.deepEqual(parseModels({ data: [
    { id: " model-a " }, { id: "model-a" }, { name: "missing-id" }, null, { id: "model-b" }
  ] }), ["model-a", "model-b"]);
  assert.throws(() => parseModels({ models: [] }), /格式无效/);
});

test("模型服务 451 使用指定中文提示", () => {
  assert.equal(modelServiceError(451), "当前网络无法访问模型服务(451)");
});

test("模型列表遇到 401 只刷新一次凭证", async () => {
  const credentials = [];
  let requests = 0;
  const models = await listModelsWithRefresh(async (force) => {
    credentials.push(force);
    return force ? "refreshed-key" : "expired-key";
  }, async (credential) => {
    requests += 1;
    if (credential === "expired-key") throw Object.assign(new Error("unauthorized"), { status: 401 });
    return ["model-a"];
  });
  assert.deepEqual(models, ["model-a"]);
  assert.deepEqual(credentials, [false, true]);
  assert.equal(requests, 2);
});

test("非 401 模型列表错误不刷新凭证", async () => {
  let credentialReads = 0;
  await assert.rejects(listModelsWithRefresh(async () => {
    credentialReads += 1;
    return "key";
  }, async () => { throw Object.assign(new Error("network"), { status: 451 }); }), /network/);
  assert.equal(credentialReads, 1);
});

test("DSH 子进程只继承允许的环境变量且用户 key 不进入环境", () => {
  const env = buildRuntimeEnv({
    baseEnv: { PATH: "allowed-path", SystemRoot: "C:\\Windows", SECRET: "not-allowed" },
    home: "C:\\Temp\\qo-run",
    gatewayUrl: "http://127.0.0.1:12345/v1",
    gatewayKey: "temporary-loopback-key",
    model: "model-a",
    system: "企业文书编辑"
  });
  assert.equal(env.HOME, env.USERPROFILE);
  assert.equal(env.DSH_HOME, env.DSH_AGENTS_HOME);
  assert.equal(env.NO_PROXY, "127.0.0.1,localhost");
  assert.equal(env.QO_RUNTIME_KEY, "temporary-loopback-key");
  assert.equal(Object.hasOwn(env, "SECRET"), false);
  assert.deepEqual(Object.keys(env).filter((key) => PASSTHROUGH_ENV.includes(key)), ["PATH", "SystemRoot"]);
  assert.deepEqual(buildDshArgs({ dshBin: "dsh.js", patchPath: "office.patch.yml" }),
    ["dsh.js", "--profile", "sdk-minimal", "--patch", "office.patch.yml"]);
});

test("开发态与安装态分别从 Harness 源目录和 resources/qo-runtime 解析路径", () => {
  const appDirectory = path.resolve("desktop");
  const development = resolveRuntimePaths(undefined, appDirectory);
  assert.equal(development.dshBin, path.resolve("harness/node_modules/@deepseek-ai/dsh/lib/bin.js"));
  const packaged = resolveRuntimePaths("C:\\App\\resources", appDirectory);
  assert.equal(packaged.dshBin, "C:\\App\\resources\\qo-runtime\\harness\\node_modules\\@deepseek-ai\\dsh\\lib\\bin.js");
  assert.equal(packaged.bundledNode, "C:\\App\\resources\\qo-runtime\\node\\node.exe");
});

test("session.fetch 显式携带工作台 qo_session cookie", async () => {
  let captured;
  const session = {
    cookies: { async get(query) {
      assert.deepEqual(query, { url: "https://office.example", name: "qo_session" });
      return [{ name: "qo_session", value: "s".repeat(43) }];
    } },
    async fetch(url, options) {
      captured = { url, options };
      return { status: 200 };
    }
  };
  await sessionFetch(session, "https://office.example", "https://office.example/api/v1/me", { method: "GET" });
  assert.equal(captured.options.headers.get("cookie"), `qo_session=${"s".repeat(43)}`);
  assert.equal(captured.options.credentials, "include");
});

test("没有 qo_session cookie 时不发起 session.fetch", async () => {
  let requested = false;
  const session = {
    cookies: { async get() { return []; } },
    async fetch() { requested = true; }
  };
  await assert.rejects(sessionFetch(session, "https://office.example", "https://office.example/api/v1/me"), /请先登录/);
  assert.equal(requested, false);
});
