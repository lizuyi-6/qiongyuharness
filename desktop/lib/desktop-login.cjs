const http = require("node:http");

const COMPLETION_HTML = "<!doctype html><html lang=\"zh-CN\"><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><title>穹驭办公</title><body>登录完成，可关闭此页。</body></html>";

function startLoopbackCallback({ timeoutMs = 120_000 } = {}) {
  let resolveCode;
  let rejectCode;
  let settled = false;
  const codePromise = new Promise((resolve, reject) => {
    resolveCode = resolve;
    rejectCode = reject;
  });
  codePromise.catch(() => {});
  const server = http.createServer((request, response) => {
    const peer = request.socket.remoteAddress;
    if (peer !== "127.0.0.1" && peer !== "::ffff:127.0.0.1") {
      response.writeHead(403).end();
      return;
    }
    let url;
    try { url = new URL(request.url, "http://127.0.0.1"); } catch {
      response.writeHead(400).end();
      return;
    }
    const code = url.searchParams.get("code");
    if (request.method !== "GET" || url.pathname !== "/callback" || !code ||
        !/^[A-Za-z0-9_-]{43,256}$/.test(code) || url.searchParams.getAll("code").length !== 1 || settled) {
      response.writeHead(404, { "Cache-Control": "no-store" }).end();
      return;
    }
    settled = true;
    response.writeHead(200, {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-store",
      "Content-Security-Policy": "default-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'",
      "X-Content-Type-Options": "nosniff",
      "Referrer-Policy": "no-referrer"
    }).end(COMPLETION_HTML);
    resolveCode(code);
    clearTimeout(timer);
    server.close();
  });

  const timer = setTimeout(() => {
    if (settled) return;
    settled = true;
    const error = new Error("本次企业登录已超时");
    rejectCode(error);
    server.close();
  }, timeoutMs);
  timer.unref?.();

  const listening = new Promise((resolve, reject) => {
    const onError = (error) => { server.off("listening", onListening); clearTimeout(timer); reject(error); };
    const onListening = () => { server.off("error", onError); resolve(server.address().port); };
    server.once("error", onError);
    server.once("listening", onListening);
    server.listen(0, "127.0.0.1");
  });

  return {
    async port() { return listening; },
    code: codePromise,
    async close() {
      clearTimeout(timer);
      if (server.listening) await new Promise((resolve) => server.close(() => resolve()));
      if (!settled) {
        settled = true;
        rejectCode(new Error("登录流程已取消"));
      }
    }
  };
}

module.exports = { COMPLETION_HTML, startLoopbackCallback };
