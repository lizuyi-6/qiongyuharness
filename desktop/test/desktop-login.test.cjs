const test = require("node:test");
const assert = require("node:assert/strict");
const { startLoopbackCallback, COMPLETION_HTML } = require("../lib/desktop-login.cjs");

test("loopback 回调仅接收一次性授权码并返回可关闭提示", async () => {
  const receiver = startLoopbackCallback({ timeoutMs: 5_000 });
  try {
    const port = await receiver.port();
    assert.ok(port >= 1024 && port <= 65535);
    const invalid = await fetch(`http://127.0.0.1:${port}/other`);
    assert.equal(invalid.status, 404);

    const code = "A".repeat(43);
    const response = await fetch(`http://127.0.0.1:${port}/callback?code=${code}`);
    assert.equal(response.status, 200);
    assert.match(response.headers.get("content-type"), /text\/html/);
    assert.equal(await response.text(), COMPLETION_HTML);
    assert.equal(await receiver.code, code);
    await assert.rejects(fetch(`http://127.0.0.1:${port}/callback?code=${code}`));
  } finally {
    await receiver.close();
  }
});
