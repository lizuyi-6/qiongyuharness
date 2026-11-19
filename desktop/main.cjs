const { app, BrowserWindow, ipcMain, shell } = require("electron");
const { createHash, randomBytes } = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

app.setName("穹驭办公");
let window;
let origin;
let loginInProgress = false;
const connectPage = pathToFileURL(path.join(__dirname, "connect.html")).href;

function normalize(value) {
  const address = new URL(value?.trim());
  const loopback = ["127.0.0.1", "localhost", "[::1]"].includes(address.hostname);
  if ((address.protocol !== "https:" && !(loopback && address.protocol === "http:")) ||
      address.username || address.password || address.search || address.hash || address.pathname !== "/") {
    throw new Error("请填写 HTTPS 服务地址。演示服务可使用 http://127.0.0.1:8765。");
  }
  return address.origin;
}

function configPath() {
  return path.join(app.getPath("userData"), "connection.json");
}

function config() {
  const command = process.argv.find((value) => value.startsWith("--server="));
  const provided = command?.slice(9) || process.env.QO_DESKTOP_URL;
  if (provided) return normalize(provided);
  try {
    return normalize(JSON.parse(fs.readFileSync(configPath(), "utf8")).url);
  } catch {
    return undefined;
  }
}

function trusted(event, local = false) {
  const sender = event.senderFrame;
  if (!window || event.sender !== window.webContents || !sender || sender.parent) return false;
  if (local) return sender.url === connectPage;
  return !!origin && new URL(sender.url).origin === origin;
}

async function openWorkplace() {
  if (!origin) return window.loadFile("connect.html");
  try {
    await window.loadURL(origin);
  } catch {
    await window.loadFile("connect.html");
  }
}

ipcMain.handle("qo:connect", async (event, value) => {
  if (!trusted(event, true)) throw new Error("页面无权修改服务地址");
  origin = normalize(value);
  fs.mkdirSync(path.dirname(configPath()), { recursive: true });
  fs.writeFileSync(configPath(), JSON.stringify({ url: origin }), { mode: 0o600 });
  setImmediate(() => { void openWorkplace(); });
});

ipcMain.handle("qo:configured-url", (event) => {
  if (trusted(event, true)) return origin || "";
  return "";
});

ipcMain.handle("qo:sign-in", async (event) => {
  if (!trusted(event) || loginInProgress) throw new Error("企业登录正在进行");
  const start = origin;
  const verifier = randomBytes(48).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  loginInProgress = true;
  try {
    await shell.openExternal(start + "/auth/start?desktop=" + challenge);
    for (let attempt = 0; attempt < 150; attempt++) {
      await new Promise((resolve) => setTimeout(resolve, 1600));
      if (!window || window.isDestroyed() || origin !== start) throw new Error("登录窗口已关闭");
      let response;
      try {
        response = await fetch(start + "/api/v1/auth/desktop/claim", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ verifier }),
          redirect: "error",
          signal: AbortSignal.timeout(5000)
        });
      } catch { continue; }
      if (response.status === 404) continue;
      if (!response.ok) throw new Error("无法完成企业登录");
      const { session } = await response.json();
      if (typeof session !== "string" || !/^[A-Za-z0-9_-]{43,128}$/.test(session)) {
        throw new Error("企业登录凭证无效");
      }
      await window.webContents.session.cookies.set({
        url: start, name: "qo_session", value: session, httpOnly: true,
        secure: start.startsWith("https:"), sameSite: "lax", path: "/"
      });
      return true;
    }
    throw new Error("本次企业登录已超时");
  } finally {
    loginInProgress = false;
  }
});

app.whenReady().then(() => {
  try { origin = config(); } catch { origin = undefined; }
  window = new BrowserWindow({
    width: 1420, height: 960, minWidth: 780, minHeight: 570,
    title: "穹驭办公", backgroundColor: "#f4f6fa", autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true, nodeIntegration: false, sandbox: true,
      webSecurity: true, partition: "qiongyu-office"
    }
  });
  const content = window.webContents;
  content.session.setPermissionRequestHandler((_sender, _permission, response) => response(false));
  content.session.setPermissionCheckHandler(() => false);
  const navigation = (event, details) => {
    const url = typeof details === "string" ? details : details.url;
    if (url === connectPage) return;
    try {
      if (origin && new URL(url).origin === origin) return;
    } catch {}
    event.preventDefault();
  };
  content.on("will-navigate", navigation);
  content.on("will-redirect", navigation);
  content.setWindowOpenHandler(({ url, userGesture }) => {
    try {
      if (userGesture !== false && new URL(url).protocol === "https:") void shell.openExternal(url);
    } catch {}
    return { action: "deny" };
  });
  window.on("closed", () => { window = undefined; });
  void openWorkplace();
});

app.on("window-all-closed", () => app.quit());
