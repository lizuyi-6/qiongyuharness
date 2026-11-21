const { contextBridge, ipcRenderer } = require("electron");

const versionArg = process.argv.find((arg) => arg.startsWith("--qo-version="));
const appVersion = versionArg ? versionArg.slice("--qo-version=".length).trim() : "";
const hasTitleBarOverlay = process.argv.includes("--qo-titlebar-overlay=1");

contextBridge.exposeInMainWorld("officeDesktop", {
  version: appVersion,
  hasTitleBarOverlay,
  setTheme: (mode) => ipcRenderer.invoke("qo:set-theme", mode),
  signIn: () => ipcRenderer.invoke("qo:sign-in"),
  models: () => ipcRenderer.invoke("qo:models"),
  usage: () => ipcRenderer.invoke("qo:usage"),
  runAgent: (job) => ipcRenderer.invoke("qo:run-agent", job),
  generateImage: (request) => ipcRenderer.invoke("qo:generate-image", request),
  generateTitle: (request) => ipcRenderer.invoke("qo:generate-title", request),
  cancelAgent: () => ipcRenderer.invoke("qo:cancel-agent"),
  getCheckpoint: (conversationId, userId) => ipcRenderer.invoke("qo:get-checkpoint", conversationId, userId),
  clearCheckpoint: (conversationId) => ipcRenderer.invoke("qo:clear-checkpoint", conversationId),
  collectLogs: (options) => ipcRenderer.invoke("qo:collect-logs", options || {}),
  openConnect: () => ipcRenderer.invoke("qo:open-connect"),
  getUpdateState: () => ipcRenderer.invoke("qo:update-state"),
  checkUpdate: () => ipcRenderer.invoke("qo:update-check"),
  downloadUpdate: (options) => ipcRenderer.invoke("qo:update-download", options),
  cancelUpdate: () => ipcRenderer.invoke("qo:update-cancel"),
  installUpdate: () => ipcRenderer.invoke("qo:update-install"),
  clearUpdateNotice: () => ipcRenderer.invoke("qo:update-clear-notice"),
  onAgentEvent: (callback) => {
    if (typeof callback !== "function") throw new TypeError("callback must be a function");
    const listener = (_event, value) => callback(value);
    ipcRenderer.on("qo:agent-event", listener);
    return () => ipcRenderer.removeListener("qo:agent-event", listener);
  },
  onUpdateEvent: (callback) => {
    if (typeof callback !== "function") throw new TypeError("callback must be a function");
    const listener = (_event, value) => callback(value);
    ipcRenderer.on("qo:update-event", listener);
    return () => ipcRenderer.removeListener("qo:update-event", listener);
  }
});

window.addEventListener("DOMContentLoaded", async () => {
  if (window.location.protocol !== "file:") return;
  const syncConnectTheme = () => {
    const dark = window.matchMedia?.("(prefers-color-scheme: dark)")?.matches === true;
    void ipcRenderer.invoke("qo:set-theme", dark ? "dark" : "light").catch(() => {});
  };
  syncConnectTheme();
  window.matchMedia?.("(prefers-color-scheme: dark)")?.addEventListener?.("change", syncConnectTheme);
  const form = document.getElementById("connection");
  if (!form) return;
  const url = document.getElementById("server");
  const notice = document.getElementById("notice");
  const configuredUrl = await ipcRenderer.invoke("qo:configured-url");
  if (configuredUrl) url.value = configuredUrl;
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    notice.className = "notice";
    notice.textContent = "正在连接…";
    try {
      await ipcRenderer.invoke("qo:connect", url.value);
    } catch (error) {
      notice.className = "notice error";
      notice.textContent = error.message.replace(/^.*?Error: /, "") || "无法连接";
    }
  });
});
