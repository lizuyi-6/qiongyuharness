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

async function loadChat() {
  const element = { addEventListener() {}, querySelector: () => null, querySelectorAll: () => [] };
  globalThis.document ??= { querySelector: () => element, querySelectorAll: () => [], addEventListener() {} };
  globalThis.window ??= globalThis;
  globalThis.sessionStorage ??= memoryStorage();
  const storage = memoryStorage();
  globalThis.localStorage = storage;
  const chat = await import(pathToFileURL(path.join(REPO_ROOT, "server", "static", "chat.js")).href);
  const ui = await import(pathToFileURL(path.join(REPO_ROOT, "server", "static", "ui.js")).href);
  return { chat, ui, storage };
}

test("AI 自行生图缺省关闭，且按账号分开记忆", async () => {
  const { chat, ui, storage } = await loadChat();

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
