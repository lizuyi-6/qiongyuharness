const { app, BrowserWindow, ipcMain, Menu, nativeImage, shell, Tray } = require("electron");
const { createHash, randomBytes } = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const { startLoopbackCallback } = require("./lib/desktop-login.cjs");
const { createDesktopUpdater } = require("./lib/updater.cjs");
const { createDesktopAgent } = require("./agent.cjs");

app.setName("穹驭办公");
// 与安装器快捷方式写入的 AppUserModelID(= build.appId)一致，任务栏才会把窗口归到固定的快捷方式并用其图标。
if (process.platform === "win32") app.setAppUserModelId("com.qiongyutech.office");
app.setPath("userData", path.join(app.getPath("appData"), "QiongyuOffice"));
let window;
let tray;
let origin;
let loginInProgress = false;
let desktopAgent;
let desktopUpdater;
let updateCheckTimer;
let isQuitting = false;
let trayHintShown = false;
const appIconPath = path.join(__dirname, "office.ico");
const connectPage = pathToFileURL(path.join(__dirname, "connect.html")).href;

const singleLock = app.requestSingleInstanceLock();
if (!singleLock) {
  app.quit();
}

function showMainWindow(routeHash = "") {
  if (!window || window.isDestroyed()) return;
  if (window.isMinimized()) window.restore();
  window.show();
  window.focus();
  if (routeHash && origin) {
    try {
      const current = window.webContents.getURL();
      if (current && new URL(current).origin === origin) {
        const targetHash = routeHash.startsWith("#") ? routeHash : `#${routeHash}`;
        void window.webContents.executeJavaScript(`location.hash = ${JSON.stringify(targetHash)};`).catch(() => {});
      }
    } catch {}
  }
}

function quitApplication() {
  isQuitting = true;
  if (updateCheckTimer) {
    clearInterval(updateCheckTimer);
    updateCheckTimer = undefined;
  }
  try { desktopAgent?.cancelAgent(); } catch {}
  try { desktopUpdater?.cancelDownload(); } catch {}
  app.quit();
}

function buildTrayMenu() {
  const version = app.getVersion();
  const visible = Boolean(window && !window.isDestroyed() && window.isVisible() && !window.isMinimized());
  const updateState = desktopUpdater?.getState?.();
  const updateReady = updateState?.status === "ready";
  const updateAvailable = updateState?.status === "available";
  const updateLabel = updateReady && updateState?.latestVersion
    ? `安装新版本 v${updateState.latestVersion}…`
    : updateAvailable && updateState?.latestVersion
      ? `发现新版本 v${updateState.latestVersion}…`
      : "检查更新…";
  return Menu.buildFromTemplate([
    {
      label: visible ? "隐藏到托盘" : "打开穹驭办公",
      click: () => {
        if (!window || window.isDestroyed()) return;
        if (window.isVisible() && !window.isMinimized()) window.hide();
        else showMainWindow();
      }
    },
    { type: "separator" },
    {
      label: "新建办公会话",
      enabled: Boolean(origin),
      click: () => showMainWindow("#chat")
    },
    {
      label: "查看今日任务",
      enabled: Boolean(origin),
      click: () => showMainWindow("#today")
    },
    {
      label: "我的文件",
      enabled: Boolean(origin),
      click: () => showMainWindow("#files")
    },
    { type: "separator" },
    {
      label: "停止当前 AI 执行",
      click: () => {
        try { desktopAgent?.cancelAgent(); } catch {}
      }
    },
    {
      label: updateLabel,
      enabled: Boolean(origin),
      click: () => {
        showMainWindow("#settings");
        if (!updateReady) void desktopUpdater?.checkForUpdates({ silent: false });
      }
    },
    {
      label: "切换工作台地址…",
      click: () => {
        showMainWindow();
        try { desktopAgent?.cancelAgent(); } catch {}
        if (window && !window.isDestroyed()) void window.loadFile("connect.html");
      }
    },
    { type: "separator" },
    {
      label: `穹驭办公 v${version}`,
      enabled: false
    },
    {
      label: "退出穹驭办公",
      click: () => quitApplication()
    }
  ]);
}

function ensureTray() {
  if (tray) return tray;
  const icon = nativeImage.createFromPath(appIconPath);
  tray = new Tray(icon);
  tray.setToolTip(`穹驭办公 v${app.getVersion()}`);
  tray.on("click", () => showMainWindow());
  tray.on("double-click", () => showMainWindow());
  tray.on("right-click", () => {
    tray.popUpContextMenu(buildTrayMenu());
  });
  return tray;
}

app.on("second-instance", () => {
  showMainWindow();
});

app.on("before-quit", () => {
  isQuitting = true;
  if (tray && !tray.isDestroyed()) {
    tray.destroy();
    tray = undefined;
  }
});

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
  const previous = origin;
  origin = normalize(value);
  if (previous && previous !== origin) await desktopAgent?.clearCredentialsForOrigin(previous);
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
  const callback = startLoopbackCallback();
  loginInProgress = true;
  try {
    const port = await callback.port();
    const loginUrl = new URL("/auth/start", start);
    loginUrl.searchParams.set("desktop", challenge);
    loginUrl.searchParams.set("port", String(port));
    await shell.openExternal(loginUrl.toString());

    const code = await callback.code;
    if (!window || window.isDestroyed() || origin !== start) throw new Error("登录窗口已关闭");
    const response = await fetch(`${start}/api/v1/auth/desktop/claim`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ code, verifier }),
      redirect: "error",
      signal: AbortSignal.timeout(15_000)
    });
    let payload = {};
    try { payload = await response.json(); } catch {}
    if (!response.ok) throw new Error(payload.detail || "无法完成企业登录，请重新登录");
    if (typeof payload.session !== "string" || !/^[A-Za-z0-9_-]{43,128}$/.test(payload.session)) {
      throw new Error("企业登录凭证无效");
    }

    await desktopAgent?.clearCredentialsForOrigin(start);
    await window.webContents.session.cookies.set({
      url: start, name: "qo_session", value: payload.session, httpOnly: true,
      secure: start.startsWith("https:"), sameSite: "lax", path: "/"
    });
    try {
      await desktopAgent?.prefetchAfterLogin();
    } catch (error) {
      window?.webContents.send("qo:agent-event", { type: "step", label: error.message || "模型列表加载失败" });
    }
    showMainWindow();
    return true;
  } finally {
    loginInProgress = false;
    await callback.close();
  }
});

ipcMain.handle("qo:models", async (event) => {
  if (!trusted(event)) throw new Error("页面无权读取模型列表");
  return desktopAgent.models();
});

ipcMain.handle("qo:usage", async (event) => {
  if (!trusted(event)) throw new Error("页面无权读取模型额度");
  return desktopAgent.usage();
});

ipcMain.handle("qo:run-agent", async (event, job) => {
  if (!trusted(event)) throw new Error("页面无权启动桌面执行");
  return desktopAgent.runAgent(job);
});

ipcMain.handle("qo:generate-image", async (event, request) => {
  if (!trusted(event)) throw new Error("页面无权生成图片");
  return desktopAgent.generateImage(request);
});

ipcMain.handle("qo:cancel-agent", (event) => {
  if (!trusted(event)) throw new Error("页面无权停止桌面执行");
  return desktopAgent.cancelAgent();
});

ipcMain.handle("qo:get-checkpoint", async (event, conversationId, userId) => {
  if (!trusted(event)) throw new Error("页面无权读取任务断点");
  return desktopAgent.getCheckpoint(conversationId, userId);
});

ipcMain.handle("qo:clear-checkpoint", async (event, conversationId) => {
  if (!trusted(event)) throw new Error("页面无权清除任务断点");
  return desktopAgent.clearCheckpoint(conversationId);
});

ipcMain.handle("qo:update-state", (event) => {
  if (!trusted(event)) throw new Error("页面无权读取更新状态");
  return desktopUpdater?.getState() || null;
});

ipcMain.handle("qo:update-check", async (event) => {
  if (!trusted(event)) throw new Error("页面无权检查更新");
  return desktopUpdater.checkForUpdates({ silent: false });
});

ipcMain.handle("qo:update-download", async (event, options) => {
  if (!trusted(event)) throw new Error("页面无权下载更新");
  return desktopUpdater.downloadUpdate(options || {});
});

ipcMain.handle("qo:update-cancel", (event) => {
  if (!trusted(event)) throw new Error("页面无权取消下载");
  return desktopUpdater.cancelDownload();
});

ipcMain.handle("qo:update-install", async (event) => {
  if (!trusted(event)) throw new Error("页面无权安装更新");
  return desktopUpdater.installUpdate();
});

ipcMain.handle("qo:update-clear-notice", (event) => {
  if (!trusted(event)) throw new Error("页面无权更新提示状态");
  return desktopUpdater.clearAppliedNotice();
});

ipcMain.handle("qo:open-connect", async (event) => {
  if (!trusted(event)) throw new Error("页面无权打开连接设置");
  desktopAgent?.cancelAgent();
  if (window && !window.isDestroyed()) await window.loadFile("connect.html");
  return true;
});

ipcMain.handle("qo:set-theme", (event, mode) => {
  if (!trusted(event) && !trusted(event, true)) return false;
  if (!window || window.isDestroyed()) return false;
  const dark = mode === "dark";
  window.setBackgroundColor(dark ? "#0b0b0c" : "#ffffff");
  if (typeof window.setTitleBarOverlay === "function") {
    window.setTitleBarOverlay({
      color: "#00000000",
      symbolColor: dark ? "#f4f4f5" : "#111113",
      height: 44
    });
  }
  return true;
});

app.whenReady().then(() => {
  if (!singleLock) return;
  try { origin = config(); } catch { origin = undefined; }
  ensureTray();
  const useTitleBarOverlay = process.platform === "win32" || process.platform === "linux";
  window = new BrowserWindow({
    width: 1420, height: 960, minWidth: 780, minHeight: 570,
    title: "穹驭办公", backgroundColor: "#ffffff", autoHideMenuBar: true,
    icon: path.join(__dirname, "office.ico"),
    titleBarStyle: "hidden",
    ...(useTitleBarOverlay ? {
      titleBarOverlay: {
        color: "#00000000",
        symbolColor: "#111113",
        height: 44
      }
    } : {}),
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      additionalArguments: [
        `--qo-version=${app.getVersion()}`,
        ...(useTitleBarOverlay ? ["--qo-titlebar-overlay=1"] : [])
      ],
      contextIsolation: true, nodeIntegration: false, sandbox: true,
      webSecurity: true, partition: "qiongyu-office"
    }
  });
  desktopAgent = createDesktopAgent({
    app,
    getWindow: () => window,
    getOrigin: () => origin,
    emit: (value) => {
      if (window && !window.isDestroyed()) window.webContents.send("qo:agent-event", value);
    }
  });
  desktopUpdater = createDesktopUpdater({
    app,
    getWindow: () => window,
    getOrigin: () => origin,
    quitApplication,
    emit: (value) => {
      if (window && !window.isDestroyed()) window.webContents.send("qo:update-event", value);
    }
  });
  const content = window.webContents;
  content.on("did-finish-load", () => {
    if (!origin) return;
    try {
      if (new URL(content.getURL()).origin === origin) {
        setTimeout(() => { void desktopUpdater?.checkForUpdates({ silent: true }); }, 4_000);
      }
    } catch {}
  });
  updateCheckTimer = setInterval(() => {
    if (origin) void desktopUpdater?.checkForUpdates({ silent: true });
  }, 4 * 60 * 60 * 1000);
  updateCheckTimer.unref?.();
  content.session.cookies.on("changed", (_event, cookie, _cause, removed) => {
    if (cookie.name === "qo_session" && removed) void desktopAgent.clearCredentialsForOrigin(origin);
  });
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
  window.on("close", (event) => {
    if (!isQuitting) {
      event.preventDefault();
      window.hide();
      if (process.platform === "win32" && tray && !trayHintShown) {
        trayHintShown = true;
        try {
          tray.displayBalloon({
            title: "穹驭办公已最小化到系统托盘",
            content: "应用仍在后台运行，进行中的任务不会中断。点击托盘图标可重新打开，右键可退出。",
            iconType: "none"
          });
        } catch {}
      }
    }
  });
  window.on("closed", () => { window = undefined; });
  void openWorkplace();
});

app.on("activate", () => showMainWindow());

app.on("window-all-closed", () => {
  if (isQuitting) app.quit();
});
