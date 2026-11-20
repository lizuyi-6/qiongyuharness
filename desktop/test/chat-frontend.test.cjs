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

test("AI 自行生图缺省关闭，且按账号分开记忆", async () => {
  const { chat, ui } = await modules;
  const storage = memoryStorage();
  globalThis.localStorage = storage;

  ui.state.me = null;
  assert.equal(chat.imageAutoPreference(), false);
  chat.saveImageAutoPreference(true);
  assert.deepEqual(storage.keys(), [], "未登录时不写入偏好");

  ui.state.me = { id: "user-a" };
  assert.equal(chat.imageAutoPreference(), false, "从未勾选过的账号缺省关闭");
  chat.saveImageAutoPreference(true);
  assert.equal(chat.imageAutoPreference(), true);

  ui.state.me = { id: "user-b" };
  assert.equal(chat.imageAutoPreference(), false, "其他账号不沿用 user-a 的勾选");

  ui.state.me = { id: "user-a" };
  chat.saveImageAutoPreference(false);
  assert.equal(chat.imageAutoPreference(), false);

  globalThis.localStorage = { getItem() { throw new Error("blocked"); }, setItem() { throw new Error("blocked"); } };
  assert.equal(chat.imageAutoPreference(), false, "存储不可用时缺省关闭");
  assert.doesNotThrow(() => chat.saveImageAutoPreference(true));
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
