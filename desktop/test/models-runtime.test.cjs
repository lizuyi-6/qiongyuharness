const test = require("node:test");
const assert = require("node:assert/strict");
const { listModelsWithRefresh, parseModels, modelServiceError } = require("../lib/models.cjs");
const { buildDshArgs, buildRuntimeEnv, PASSTHROUGH_ENV, resolveRuntimePaths } = require("../lib/runtime-env.cjs");
const path = require("node:path");
const { sessionFetch } = require("../lib/session-cookie.cjs");
const { quotaExhaustionMessage, requestUsage, sanitizeUsage } = require("../lib/usage.cjs");

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

test("额度净化保留 quota_limited 的 5 小时和周字段", () => {
  const usage = sanitizeUsage({
    mode: "quota_limited",
    quota: { limit: 100, used: 25, remaining: 75, unit: "USD", private_note: "drop" },
    rate_limits: [
      { window: "5h", limit: 10, used: 4, remaining: 6, reset_at: "2026-11-20T09:30:00Z", extra: "drop" },
      { window: "7d", limit: 50, used: 20, remaining: 30, reset_at: "2026-11-23T00:00:00+08:00" }
    ],
    api_key: "must-not-cross-ipc"
  });
  assert.deepEqual(usage.rate_limits.map(({ window, limit, used, remaining, reset_at }) =>
    ({ window, limit, used, remaining, reset_at })), [
    { window: "5h", limit: 10, used: 4, remaining: 6, reset_at: "2026-11-20T09:30:00Z" },
    { window: "7d", limit: 50, used: 20, remaining: 30, reset_at: "2026-11-23T00:00:00+08:00" }
  ]);
  assert.deepEqual(usage.quota, { limit: 100, used: 25, remaining: 75, unit: "USD" });
  assert.equal(Object.hasOwn(usage, "api_key"), false);
  assert.equal(Object.hasOwn(usage.rate_limits[0], "extra"), false);
});

test("额度净化保留订阅周额度和起始时间", () => {
  const usage = sanitizeUsage({
    mode: "unrestricted",
    planName: "团队套餐",
    subscription: {
      daily_usage_usd: 2, weekly_usage_usd: 12, monthly_usage_usd: 32,
      daily_limit_usd: 10, weekly_limit_usd: 80, monthly_limit_usd: 250,
      weekly_window_start: "2026-11-16T00:00:00+08:00", expires_at: "2026-12-01T00:00:00Z",
      internal_id: 7
    }
  });
  assert.equal(usage.subscription.weekly_usage_usd, 12);
  assert.equal(usage.subscription.weekly_limit_usd, 80);
  assert.equal(usage.subscription.weekly_window_start, "2026-11-16T00:00:00+08:00");
  assert.equal(Object.hasOwn(usage.subscription, "internal_id"), false);
});

test("额度净化保留钱包余额并丢弃未知字段", () => {
  const usage = sanitizeUsage({ mode: "unrestricted", planName: "钱包余额", balance: 23.5, remaining: 23.5, hidden: true });
  assert.deepEqual(usage, { mode: "unrestricted", planName: "钱包余额", balance: 23.5, remaining: 23.5 });
});

test("额度净化拒绝字符串数字和布尔数字", () => {
  const stringNumber = sanitizeUsage({ mode: "quota_limited", rate_limits: [
    { window: "5h", limit: "5", used: 1, remaining: 4 }
  ] });
  const booleanNumber = sanitizeUsage({ mode: "quota_limited", rate_limits: [
    { window: "5h", limit: 5, used: true, remaining: 4 }
  ] });
  assert.equal(stringNumber, null);
  assert.equal(booleanNumber, null);
});

test("额度净化拒绝未支持的窗口和无效时间", () => {
  assert.equal(sanitizeUsage({ mode: "quota_limited", rate_limits: [
    { window: "30d", limit: 5, used: 1, remaining: 4 }
  ] }), null);
  assert.equal(sanitizeUsage({ mode: "quota_limited", rate_limits: [
    { window: "5h", limit: 5, used: 1, remaining: 4, reset_at: "tomorrow" }
  ] }), null);
  assert.equal(sanitizeUsage({ mode: "quota_limited", rate_limits: Array.from({ length: 9 }, () =>
    ({ window: "5h", limit: 5, used: 1, remaining: 4 })) }), null);
  assert.equal(sanitizeUsage({ mode: "quota_limited", rate_limits: [
    { window: "5h", limit: 5, used: Number.NaN, remaining: 4 }
  ] }), null);
});

test("额度查询把 401 和 403 映射为 unauthorized", async () => {
  for (const status of [401, 403]) {
    const result = await requestUsage({ baseUrl: "https://usage.example/v1", apiKey: "test-key" }, async (url, options) => {
      assert.equal(url, "https://usage.example/v1/usage");
      assert.equal(options.headers.Authorization, "Bearer test-key");
      return { status, ok: false };
    });
    assert.deepEqual(result, { state: "unauthorized" });
  }
});

test("额度查询把 5xx 和网络错误映射为 unavailable", async () => {
  const credential = { baseUrl: "https://usage.example/v1", apiKey: "test-key" };
  assert.deepEqual(await requestUsage(credential, async () => ({ status: 503, ok: false })), { state: "unavailable" });
  assert.deepEqual(await requestUsage(credential, async () => { throw new Error("offline"); }), { state: "unavailable" });
});

test("额度耗尽错误显示中文窗口和本地化重置时间", () => {
  const message = quotaExhaustionMessage({
    type: "rate_limit_exceeded", code: "USER_PLATFORM_WEEKLY_QUOTA_EXHAUSTED",
    resetAt: "2026-11-20T09:30:00Z"
  }, null, { timeZone: "UTC" });
  assert.match(message, /额度已用尽（周）/);
  assert.match(message, /将于 2026年11月20日 09:30:00 恢复/);
});
