const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const Module = require("node:module");
const { listModelsWithRefresh, modelServiceCode, parseModels, modelServiceError } = require("../lib/models.cjs");
const { buildDshArgs, buildRuntimeEnv, PASSTHROUGH_ENV, resolveRuntimePaths } = require("../lib/runtime-env.cjs");
const { sessionFetch } = require("../lib/session-cookie.cjs");
const { quotaExhaustionMessage, requestUsage, sanitizeUsage } = require("../lib/usage.cjs");

function loadAgent(netFetch) {
  const originalLoad = Module._load;
  Module._load = function mockElectron(request, parent, isMain) {
    if (request === "electron") {
      return {
        net: { fetch: netFetch },
        safeStorage: {
          isEncryptionAvailable: () => true,
          encryptString: (value) => Buffer.from(value),
          decryptString: (value) => value.toString()
        }
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

function makeModelsAgent(netFetch) {
  const { createDesktopAgent } = loadAgent(netFetch);
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), "qo-models-test-"));
  const session = {
    cookies: { get: async () => [{ name: "qo_session", value: "s".repeat(43) }] },
    fetch: async (url) => {
      let payload;
      if (String(url).endsWith("/api/v1/me")) payload = { id: "u1" };
      else if (String(url).endsWith("/api/v1/me/model-credential")) payload = {
        api_key: "test-key", key_id: 1, group_id: 7,
        base_url: "https://models.example/v1"
      };
      else throw new Error("unexpected workbench request");
      return { ok: true, status: 200, text: async () => JSON.stringify(payload) };
    }
  };
  const agent = createDesktopAgent({
    app: { getPath: () => userData, isPackaged: false },
    getWindow: () => ({ isDestroyed: () => false, webContents: { session } }),
    getOrigin: () => "https://office.example",
    dependencies: {
      fetchSessionJson: async () => { throw new Error("pipeline endpoint intentionally unavailable"); }
    }
  });
  return { agent, cleanup: () => fs.rmSync(userData, { recursive: true, force: true }) };
}

test("模型列表解析只返回去重后的模型 id", () => {
  assert.deepEqual(parseModels({ data: [
    { id: " model-a " }, { id: "model-a" }, { name: "missing-id" }, null, { id: "model-b" }
  ] }), ["model-a", "model-b"]);
  assert.throws(() => parseModels({ models: [] }), /格式无效/);
});

test("模型服务 451 使用指定中文提示", () => {
  assert.equal(modelServiceError(451), "当前网络无法访问模型服务(451)");
});

test("模型服务错误体按状态和错误码映射为明确中文提示", () => {
  const cases = [
    [401, {}, "模型凭证无效，请重新登录后重试"],
    [451, {}, "当前网络无法访问模型服务(451)"],
    [403, { code: "INSUFFICIENT_BALANCE", message: "balance zero" }, "模型账户余额不足，请联系管理员开通穹驭套餐"],
    [403, { code: "SUBSCRIPTION_NOT_FOUND" }, "尚未开通穹驭套餐或套餐已失效，请联系管理员"],
    [403, { error: { code: "SUBSCRIPTION_INVALID", type: "invalid_request", message: "expired" } }, "尚未开通穹驭套餐或套餐已失效，请联系管理员"],
    [403, { code: "API_KEY_EXPIRED" }, "模型凭证已过期，请重新登录后重试"],
    [403, { error: { code: "ACCESS_DENIED" } }, "当前网络不允许访问模型服务(403)"],
    [403, { code: "GROUP_DISABLED" }, "穹驭分组当前不可用，请联系管理员"],
    [403, { code: "GROUP_DELETED" }, "穹驭分组当前不可用，请联系管理员"],
    [403, { code: "GROUP_NOT_ALLOWED" }, "穹驭分组当前不可用，请联系管理员"],
    [429, {}, "模型额度已用完，请稍后再试"]
  ];
  for (const [status, payload, expected] of cases) {
    assert.equal(modelServiceError(status, payload), expected);
  }
  const generic = modelServiceError(500, { detail: "Backend\nfailed" });
  assert.equal(generic, "模型服务返回错误(HTTP 500): Backendfailed");
  assert.doesNotMatch(generic, /检查网络连接/);
  assert.equal(modelServiceError(502, { message: "x".repeat(205) }).length,
    "模型服务返回错误(HTTP 502): ".length + 200);
  assert.equal(modelServiceCode({ error: { code: "SUBSCRIPTION_NOT_FOUND" } }), "SUBSCRIPTION_NOT_FOUND");
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

test("模型列表遇到需补发订阅的 403 只刷新一次凭证", async () => {
  for (const code of ["SUBSCRIPTION_NOT_FOUND", "INSUFFICIENT_BALANCE"]) {
    const credentials = [];
    let requests = 0;
    await listModelsWithRefresh(async (force) => {
      credentials.push(force);
      return force ? "refreshed-key" : "old-key";
    }, async (credential) => {
      requests += 1;
      if (credential === "old-key") throw Object.assign(new Error(code), { status: 403, code });
      return ["model-a"];
    });
    assert.deepEqual(credentials, [false, true]);
    assert.equal(requests, 2);
  }
});

test("非 401 模型列表错误不刷新凭证", async () => {
  let credentialReads = 0;
  await assert.rejects(listModelsWithRefresh(async () => {
    credentialReads += 1;
    return "key";
  }, async () => { throw Object.assign(new Error("network"), { status: 451 }); }), /network/);
  assert.equal(credentialReads, 1);
});

test("403 ACCESS_DENIED 不刷新凭证", async () => {
  const credentials = [];
  let requests = 0;
  await assert.rejects(listModelsWithRefresh(async (force) => {
    credentials.push(force);
    return "key";
  }, async () => {
    requests += 1;
    throw Object.assign(new Error("denied"), { status: 403, code: "ACCESS_DENIED" });
  }), /denied/);
  assert.deepEqual(credentials, [false]);
  assert.equal(requests, 1);
});

test("模型列表重试失败时原样抛出第二次错误且不进行第三次请求", async () => {
  const credentials = [];
  let requests = 0;
  let subscriptionCreated = false;
  const retryError = Object.assign(new Error("retry failed"), { status: 403, code: "SUBSCRIPTION_NOT_FOUND" });
  await assert.rejects(listModelsWithRefresh(async (force) => {
    credentials.push(force);
    if (force) subscriptionCreated = true;
    return "key";
  }, async () => {
    requests += 1;
    if (requests === 1) throw Object.assign(new Error("refresh me"), { status: 403, code: "INSUFFICIENT_BALANCE" });
    throw retryError;
  }), (error) => error === retryError);
  assert.deepEqual(credentials, [false, true]);
  assert.equal(requests, 2);
  assert.equal(subscriptionCreated, true);
});

test("凭证强制刷新在外部开通生效前失败时不发第二次模型请求", async () => {
  const credentials = [];
  let requests = 0;
  let subscriptionCreated = false;
  const provisionError = new Error("provision failed before creating subscription");
  await assert.rejects(listModelsWithRefresh(async (force) => {
    credentials.push(force);
    if (force) throw provisionError;
    return "old-key";
  }, async () => {
    requests += 1;
    throw Object.assign(new Error("no subscription"), { status: 403, code: "SUBSCRIPTION_NOT_FOUND" });
  }), (error) => error === provisionError);
  assert.deepEqual(credentials, [false, true]);
  assert.equal(requests, 1);
  assert.equal(subscriptionCreated, false);
});

test("模型服务网络错误与超时分别显示连接和超时文案", async (t) => {
  for (const scenario of [
    { name: "connection failure", error: new TypeError("fetch failed"), expected: "无法连接模型服务，请检查网络连接后重试" },
    { name: "timeout", error: Object.assign(new Error("timed out"), { name: "TimeoutError" }), expected: "模型服务响应超时，请稍后重试" }
  ]) await t.test(scenario.name, async () => {
    const { agent, cleanup } = makeModelsAgent(async () => { throw scenario.error; });
    try {
      const result = await agent.models();
      assert.equal(result.error, scenario.expected);
    } finally {
      cleanup();
    }
  });
});

test("模型服务的余额 403 错误码被挂到错误并触发一次凭证刷新", async () => {
  let requests = 0;
  const { agent, cleanup } = makeModelsAgent(async () => {
    requests += 1;
    if (requests === 1) return {
      ok: false, status: 403,
      json: async () => ({ code: "INSUFFICIENT_BALANCE", message: "Insufficient account balance" })
    };
    return { ok: true, json: async () => ({ data: [{ id: "model-a" }] }) };
  });
  try {
    const result = await agent.models();
    assert.deepEqual(result.models, ["model-a"]);
    assert.equal(requests, 2);
  } finally {
    cleanup();
  }
});

test("工作台请求错误使用 detail、登录失效和 HTTP 状态独立文案", () => {
  const { responseError } = loadAgent(async () => { throw new Error("unexpected fetch"); });
  assert.equal(responseError({ detail: "无权访问" }, 403), "无权访问");
  assert.equal(responseError({}, 401), "登录已失效，请重新登录");
  assert.equal(responseError({}, 500), "工作台请求失败(HTTP 500)");
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
