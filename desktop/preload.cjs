const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("officeDesktop", {
  signIn: () => ipcRenderer.invoke("qo:sign-in"),
  models: () => ipcRenderer.invoke("qo:models"),
  usage: () => ipcRenderer.invoke("qo:usage"),
  runAgent: (job) => ipcRenderer.invoke("qo:run-agent", job),
  cancelAgent: () => ipcRenderer.invoke("qo:cancel-agent"),
  openConnect: () => ipcRenderer.invoke("qo:open-connect"),
  onAgentEvent: (callback) => {
    if (typeof callback !== "function") throw new TypeError("callback must be a function");
    const listener = (_event, value) => callback(value);
    ipcRenderer.on("qo:agent-event", listener);
    return () => ipcRenderer.removeListener("qo:agent-event", listener);
  }
});

window.addEventListener("DOMContentLoaded", async () => {
  if (window.location.protocol !== "file:") return;
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
