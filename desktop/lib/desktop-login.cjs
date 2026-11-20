const http = require("node:http");

const COMPLETION_HTML = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>登录完成 · 穹驭办公</title>
<style>
:root{
  --rail:#f7f7f8;
  --page:#ffffff;
  --subtle:#fafafa;
  --ink:#111113;
  --ink-2:#52525b;
  --ink-3:#71717a;
  --line:#e8e8ec;
  --accent:#18181b;
  --inverse:#ffffff;
  --hairline:#18181b;
  --ok:#15803d;
  --ok-bg:#f0fdf4;
  --ok-border:#bbf7d0;
  --shadow:0 20px 48px rgba(9,9,11,.08),0 2px 10px rgba(9,9,11,.04);
}
@media (prefers-color-scheme:dark){
  :root{
    --rail:#111113;
    --page:#151518;
    --subtle:#1b1b1f;
    --ink:#f4f4f5;
    --ink-2:#a1a1aa;
    --ink-3:#71717a;
    --line:#26262c;
    --accent:#fafafa;
    --inverse:#09090b;
    --hairline:#ffffff;
    --ok:#4ade80;
    --ok-bg:rgba(74,222,128,.1);
    --ok-border:rgba(74,222,128,.28);
    --shadow:0 24px 56px rgba(0,0,0,.52),0 0 0 1px rgba(255,255,255,.05);
  }
}
*{box-sizing:border-box}
body{
  margin:0;
  min-height:100vh;
  padding:24px;
  display:grid;
  place-items:center;
  background:var(--rail);
  color:var(--ink);
  font:14px/1.6 "PingFang SC","Hiragino Sans GB","Microsoft YaHei UI","Microsoft YaHei","Noto Sans CJK SC","Segoe UI",system-ui,sans-serif;
  -webkit-font-smoothing:antialiased;
}
main{
  position:relative;
  width:min(420px,100%);
  padding:32px 30px 26px;
  border:1px solid var(--line);
  border-radius:18px;
  background:var(--page);
  box-shadow:var(--shadow);
  text-align:center;
  animation:card-in 320ms cubic-bezier(.22,1,.36,1) both;
}
main::before{
  content:"";
  position:absolute;
  left:28px;
  right:28px;
  top:-1px;
  height:1.5px;
  border-radius:2px;
  background:var(--hairline);
  transform-origin:center;
  animation:line-in 480ms cubic-bezier(.22,1,.36,1) 60ms both;
}
@keyframes card-in{from{opacity:0;transform:translateY(10px) scale(.985)}to{opacity:1;transform:none}}
@keyframes line-in{from{transform:scaleX(0)}to{transform:scaleX(1)}}
.brand{
  display:inline-flex;
  align-items:center;
  gap:8px;
  margin-bottom:22px;
}
.brand-mark{
  display:grid;
  place-items:center;
  width:26px;
  height:26px;
  border-radius:8px;
  background:var(--accent);
  color:var(--inverse);
}
.brand-mark svg{width:18px;height:18px}
.brand-name{font-size:14px;font-weight:660;letter-spacing:.01em;color:var(--ink)}
.status-badge{
  margin:0 auto 16px;
  width:52px;
  height:52px;
  border-radius:50%;
  display:grid;
  place-items:center;
  background:var(--ok-bg);
  border:1px solid var(--ok-border);
  color:var(--ok);
}
.status-badge svg{
  width:24px;
  height:24px;
  stroke-dasharray:32;
  stroke-dashoffset:32;
  animation:check-draw 420ms cubic-bezier(.22,1,.36,1) 180ms forwards;
}
@keyframes check-draw{to{stroke-dashoffset:0}}
h1{
  margin:0 0 8px;
  font-size:19px;
  font-weight:650;
  letter-spacing:-.02em;
  color:var(--ink);
}
p{
  margin:0;
  font-size:13px;
  line-height:1.65;
  color:var(--ink-2);
}
.foot{
  margin-top:22px;
  padding-top:16px;
  border-top:1px solid var(--line);
  display:inline-flex;
  align-items:center;
  justify-content:center;
  gap:7px;
  width:100%;
  font-size:12px;
  color:var(--ink-3);
}
.dot{
  width:6px;
  height:6px;
  border-radius:50%;
  background:var(--ok);
  flex:none;
}
</style>
</head>
<body>
<main>
  <div class="brand">
    <span class="brand-mark" aria-hidden="true">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="butt" stroke-linejoin="miter">
        <path d="M13 18H6V6h12v7" stroke-width="2"/>
        <path d="M9.6 13.4V9.6h3.8" stroke-width="1.45"/>
        <path d="M12.6 12.6L18.8 18.8" stroke-width="2"/>
      </svg>
    </span>
    <span class="brand-name">穹驭办公</span>
  </div>
  <div class="status-badge" aria-hidden="true">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M5 12.5 9.8 17 19 7.5"/>
    </svg>
  </div>
  <h1>企业身份验证已完成</h1>
  <p>登录凭证已安全回传至穹驭办公桌面端，窗口将自动回到前台。<br>你现在可以关闭此浏览器标签页。</p>
  <div class="foot">
    <span class="dot" aria-hidden="true"></span>
    <span>登录已生效 · 本地安全回环通道</span>
  </div>
</main>
</body>
</html>`;

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
      "Content-Security-Policy": "default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'",
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
