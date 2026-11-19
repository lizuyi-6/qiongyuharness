const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("officeDesktop", {
  signIn: () => ipcRenderer.invoke("qo:sign-in"),
  models: () => ipcRenderer.invoke("qo:models"),
  runAgent: (job) => ipcRenderer.invoke("qo:run-agent", job),
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
  url.value = await ipcRenderer.invoke("qo:configured-url") || "";
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    notice.textContent = "正在连接…";
    try {
      await ipcRenderer.invoke("qo:connect", url.value);
    } catch (error) {
      notice.textContent = error.message.replace(/^.*?Error: /, "") || "无法连接";
    }
  });
});
