const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const REPO_ROOT = path.resolve(__dirname, "..", "..");

function memoryStorage() {
  const values = new Map();
  return {
    getItem: (key) => (values.has(key) ? values.get(key) : null),
    setItem: (key, value) => { values.set(key, String(value)); },
    removeItem: (key) => { values.delete(key); },
    keys: () => [...values.keys()]
  };
}

function stubElement() {
  return {
    textContent: "", value: "", style: {}, timer: undefined,
    classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
    addEventListener() {}, querySelector: () => null, querySelectorAll: () => [],
    append() {}, remove() { this.removed = true; }, select() { this.selected = true; }
  };
}

// 浏览器全局只在本文件里桩一次：ES module 在同一进程内只求值一次。
const toast = stubElement();
const created = [];
const page = {
  execCommand: () => false,
  querySelector: (selector) => (selector === "#toast" ? toast : stubElement()),
  querySelectorAll: () => [],
  addEventListener() {},
  createElement() { const node = stubElement(); created.push(node); return node; },
  body: { append() {} }
};
globalThis.document = page;
globalThis.window ??= globalThis;
globalThis.sessionStorage = memoryStorage();
globalThis.localStorage = memoryStorage();
let clipboard;
Object.defineProperty(globalThis, "navigator", { configurable: true, value: { get clipboard() { return clipboard; } } });

const modules = (async () => {
  const chat = await import(pathToFileURL(path.join(REPO_ROOT, "server", "static", "chat.js")).href);
  const ui = await import(pathToFileURL(path.join(REPO_ROOT, "server", "static", "ui.js")).href);
  return { chat, ui };
})();

test("AI 自行生图缺省开启，且按账号分开记忆", async () => {
  const { chat, ui } = await modules;
  const storage = memoryStorage();
  globalThis.localStorage = storage;

  ui.state.me = null;
  assert.equal(chat.imageAutoPreference(), true);
  chat.saveImageAutoPreference(false);
  assert.deepEqual(storage.keys(), [], "未登录时不写入偏好");

  ui.state.me = { id: "user-a" };
  assert.equal(chat.imageAutoPreference(), true, "从未关闭过的账号缺省开启");
  chat.saveImageAutoPreference(false);
  assert.equal(chat.imageAutoPreference(), false);

  ui.state.me = { id: "user-b" };
  assert.equal(chat.imageAutoPreference(), true, "其他账号不沿用 user-a 的关闭");

  ui.state.me = { id: "user-a" };
  chat.saveImageAutoPreference(true);
  assert.equal(chat.imageAutoPreference(), true);

  globalThis.localStorage = { getItem() { throw new Error("blocked"); }, setItem() { throw new Error("blocked"); } };
  assert.equal(chat.imageAutoPreference(), true, "存储不可用时缺省开启");
  assert.doesNotThrow(() => chat.saveImageAutoPreference(false));
});

test("保存被 4xx 拒绝时判定作废断点结果，401/409 与 5xx 不作废", async () => {
  const { chat } = await modules;
  assert.equal(chat.shouldInvalidateTurnCheckpoint(422), true, "内容校验 422 必须作废重放源");
  assert.equal(chat.shouldInvalidateTurnCheckpoint(400), true);
  assert.equal(chat.shouldInvalidateTurnCheckpoint(413), true);
  assert.equal(chat.shouldInvalidateTurnCheckpoint(401), false, "登录失效不作废");
  assert.equal(chat.shouldInvalidateTurnCheckpoint(409), false, "本轮已结束不作废");
  assert.equal(chat.shouldInvalidateTurnCheckpoint(500), false, "服务端错误保留断点");
  assert.equal(chat.shouldInvalidateTurnCheckpoint(502), false);
  assert.equal(chat.shouldInvalidateTurnCheckpoint(undefined), false);
  assert.equal(chat.shouldInvalidateTurnCheckpoint(0), false);
  assert.equal(chat.shouldInvalidateTurnCheckpoint("422"), false, "非整数状态码不作废");

  assert.equal(chat.shouldReportTurnFail("agent", 422), true, "非 complete 阶段失败需上报 /fail");
  assert.equal(chat.shouldReportTurnFail("complete", 422), false, "complete 422 已由服务端落盘 failed，不再多余 /fail 触发 409");
  assert.equal(chat.shouldReportTurnFail("complete", 502), false, "complete 502 已由服务端落盘 failed，不再多余 /fail 触发 409");
  assert.equal(chat.shouldReportTurnFail("complete", 0), true, "complete 网络断连需 /fail 兜底");
  assert.equal(chat.shouldReportTurnFail("complete", 504), true, "complete 504 未到达服务端需 /fail 兜底");
});

test("saveBilledAsset 在 retryServerErrors=false 时不对 502 重试", async (t) => {
  const { chat } = await modules;
  const originalFetch = globalThis.fetch;
  t.after(() => { globalThis.fetch = originalFetch; });

  let count = 0;
  globalThis.fetch = async () => {
    count += 1;
    return { ok: false, status: 502, json: async () => ({ detail: "模型输出不可读" }) };
  };
  await assert.rejects(
    chat.saveBilledAsset("/c", {}, { attempts: 3, backoffMs: 1, retryServerErrors: false }),
    (error) => error.status === 502
  );
  assert.equal(count, 1);
});

test("saveBilledAsset 网络层失败自动重试，4xx 不重试，重试耗尽后如实抛出", async (t) => {
  const { chat } = await modules;
  const originalFetch = globalThis.fetch;
  t.after(() => { globalThis.fetch = originalFetch; });

  await t.test("首次网络失败后第二次成功", async () => {
    const calls = [];
    globalThis.fetch = async () => {
      calls.push(1);
      if (calls.length === 1) throw new TypeError("Failed to fetch");
      return { ok: true, json: async () => ({ images: [] }) };
    };
    const result = await chat.saveBilledAsset("/conversations/c/complete", { a: 1 }, { backoffMs: 1 });
    assert.deepEqual(result, { images: [] });
    assert.equal(calls.length, 2);
  });

  await t.test("422 立即抛出不重试", async () => {
    let count = 0;
    globalThis.fetch = async () => {
      count += 1;
      return { ok: false, status: 422, json: async () => ({ detail: "内容校验未通过" }) };
    };
    await assert.rejects(chat.saveBilledAsset("/c", {}, { backoffMs: 1 }), (error) => error.status === 422);
    assert.equal(count, 1);
  });

  await t.test("连续网络失败重试耗尽后抛出网络错误", async () => {
    let count = 0;
    globalThis.fetch = async () => {
      count += 1;
      throw new TypeError("Failed to fetch");
    };
    await assert.rejects(chat.saveBilledAsset("/c", {}, { attempts: 3, backoffMs: 1 }),
      (error) => error.status === 0 && error.transient === true);
    assert.equal(count, 3);
  });
});

test("复制回复：clipboard 被拒时改用 execCommand，并按其返回值提示", async (t) => {
  const { chat, ui } = await modules;
  ui.state.chatWorkspace = { session: { turns: [{ reply: "第一段回复" }] } };

  await t.test("clipboard.writeText 成功", async () => {
    const written = [];
    let execCalls = 0;
    clipboard = { writeText: async (text) => { written.push(text); } };
    page.execCommand = () => { execCalls += 1; return true; };
    await chat.copyReply(0);
    assert.deepEqual(written, ["第一段回复"]);
    assert.equal(execCalls, 0);
    assert.equal(toast.textContent, "已复制回复内容");
  });

  await t.test("clipboard 被权限拒绝时改用 execCommand 且成功", async () => {
    const commands = [];
    clipboard = { writeText: async () => { throw Object.assign(new Error("denied"), { name: "NotAllowedError" }); } };
    page.execCommand = (command) => { commands.push(command); return true; };
    created.length = 0;
    await chat.copyReply(0);
    assert.deepEqual(commands, ["copy"]);
    assert.equal(created.at(-1).value, "第一段回复");
    assert.equal(created.at(-1).removed, true, "临时 textarea 用后移除");
    assert.equal(toast.textContent, "已复制回复内容");
  });

  await t.test("execCommand 返回 false 时提示失败而不是成功", async () => {
    clipboard = { writeText: async () => { throw new Error("denied"); } };
    page.execCommand = () => false;
    toast.textContent = "";
    await chat.copyReply(0);
    assert.equal(toast.textContent, "复制失败，请手动选择文本复制");
  });

  await t.test("没有 clipboard API 时直接走 execCommand", async () => {
    clipboard = undefined;
    page.execCommand = () => true;
    toast.textContent = "";
    await chat.copyReply(0);
    assert.equal(toast.textContent, "已复制回复内容");
  });

  await t.test("没有 clipboard API 且 execCommand 返回 false 时提示失败", async () => {
    clipboard = undefined;
    page.execCommand = () => false;
    toast.textContent = "";
    await chat.copyReply(0);
    assert.equal(toast.textContent, "复制失败，请手动选择文本复制");
  });
});
