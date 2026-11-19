export const root = document.querySelector("#app");
export const dialog = document.querySelector("#dialog");
const toast = document.querySelector("#toast");
export const state = {
  config: null, me: null, engine: null, route: [], today: [], filter: "all", focused: null,
  agent: null, chatWorkspace: null, admin: { offset: 0, department: "", status: "", date: "" }, auditOffset: 0
};
const paths = {
  mark: '<path d="M3 5h10l8 8v6h-6v-4l-5-5H3z" fill="currentColor" stroke="none"/><path d="M4 13h5l5 5v4H4z" fill="currentColor" opacity=".65" stroke="none"/>',
  message: '<path d="M5 4h14a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H9l-5 4V6a2 2 0 0 1 1-2Z"/><path d="M8 9h9m-9 4h6"/>',
  today: '<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M8 2v4m8-4v4M3 9h18m-13 6 2 2 5-5"/>',
  folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/>',
  history: '<path d="M3 12a9 9 0 1 1 3 6.7M3 16v-4h4m5-5v5l4 2"/>',
  sliders: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.6 1.6 0 0 1 .3 1.8l.1.1-1.6 1.6-.1-.1a1.6 1.6 0 0 1-1.8-.2 1.6 1.6 0 0 0-2.5 1v.2h-2.3v-.2a1.6 1.6 0 0 0-2.5-1 1.6 1.6 0 0 1-1.8.2l-.1.1L5.5 17l.1-.1a1.6 1.6 0 0 1 .2-1.8 1.6 1.6 0 0 0-1-2.5h-.2v-2.3h.2a1.6 1.6 0 0 0 1-2.5 1.6 1.6 0 0 1-.2-1.8l-.1-.1L7.1 4.3l.1.1A1.6 1.6 0 0 1 9 4.6a1.6 1.6 0 0 0 2.5-1v-.2h2.3v.2a1.6 1.6 0 0 0 2.5 1 1.6 1.6 0 0 1 1.8-.2l.1-.1 1.6 1.6-.1.1a1.6 1.6 0 0 1-.3 1.8 1.6 1.6 0 0 0 1 2.5h.2v2.3h-.2a1.6 1.6 0 0 0-1 2.4Z"/>',
  tasks: '<rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V2h6v2M8 10h8m-8 4h8m-8 4h5"/>',
  users: '<circle cx="9" cy="8" r="3"/><path d="M2 20v-2a4 4 0 0 1 4-4h6a4 4 0 0 1 4 4v2m1-15a3 3 0 0 1 0 6m5 9v-2a4 4 0 0 0-4-4"/>',
  audit: '<path d="M12 22s7-4 7-11V5l-7-3-7 3v6c0 7 7 11 7 11Z"/><path d="m9 11 2 2 4-4"/>',
  right: '<path d="m9 5 7 7-7 7"/>', down:'<path d="m5 9 7 7 7-7"/>', left:'<path d="m15 5-7 7 7 7"/>',
  arrow: '<path d="M5 12h14m-6-6 6 6-6 6"/>',
  search: '<circle cx="10.8" cy="10.8" r="6.8"/><path d="m16 16 4 4"/>',
  refresh: '<path d="M20 11a8 8 0 1 0-1.3 5m1.3-9v4h-4"/>',
  calendar: '<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M8 2v4m8-4v4M3 9h18"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
  alert: '<path d="M10.3 4.5a2 2 0 0 1 3.4 0l7.5 13A2 2 0 0 1 19.4 20H4.6a2 2 0 0 1-1.8-2.5zM12 9v4m0 3h.01"/>',
  docx: '<path d="M6 3h8l4 4v14H6zM14 3v5h4m-9 4h6m-6 4h6"/>',
  pptx: '<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M12 17v4m-3 0h6M8 13v-3m4 3V8m4 5v-2"/>',
  xlsx: '<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M4 9h16M4 14h16M10 9v11"/>',
  upload: '<path d="M12 16V4m-5 5 5-5 5 5M4 15v5h16v-5"/>',
  download: '<path d="M12 4v12m-5-5 5 5 5-5M4 16v4h16v-4"/>',
  edit: '<path d="m4 16-1 5 5-1L20 8l-4-4ZM14 6l4 4"/>',
  plus: '<path d="M12 5v14M5 12h14"/>', x:'<path d="m6 6 12 12M6 18 18 6"/>',
  logout:'<path d="M11 4H4v16h7m5-13 5 5-5 5m5-5H8"/>',
  link:'<path d="M13 5h6v6m0-6-8 8M9 5H5v14h14v-4"/>',
  clip:'<path d="m9 11 5-5a3 3 0 1 1 4 4l-8 8a4 4 0 1 1-6-6l8-8"/>'
};
export function icon(name) {
  return `<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linecap="round" stroke-linejoin="round">${paths[name] || paths.docx}</svg>`;
}

export function e(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
}

export function outputWarnings(telemetry) {
  if (!telemetry || typeof telemetry !== "object") return "";
  const messages = [];
  if (telemetry.truncated === true) messages.push("回复达到单次输出上限(3500 tokens)被截断");
  if (Array.isArray(telemetry.max_tokens_clamped)) {
    for (const item of telemetry.max_tokens_clamped) {
      if (Number.isInteger(item?.requested) && item?.applied === 3500) {
        messages.push(`请求的输出上限 ${item.requested} 已被压到 3500`);
      }
    }
  }
  return messages.length ? `<div class="output-warnings" role="status">${messages.map(e).join("<br>")}</div>` : "";
}

export function uuid() {
  if (crypto.randomUUID) return crypto.randomUUID();
  const bytes = crypto.getRandomValues(new Uint8Array(16));
  return Array.from(bytes, (v) => v.toString(16).padStart(2, "0")).join("");
}
export const isManager = () => ["admin", "manager"].includes(state.me?.role);
export const isReview = () => isManager() || state.me?.role === "reviewer";
export const active = (status) => ["pending", "in_progress", "paused", "returned"].includes(status);
export const statuses = {
  pending:"待开始",in_progress:"进行中",paused:"已暂停",returned:"退回修改",submitted:"待审核",
  approved:"已验收",closed:"已关闭",revoked:"已撤销",scheduled:"待发布",
  overdue:"已逾期",draft:"草稿",active:"运行中",finished:"已结束"
};
export const labels = {docx:"Word",pptx:"PowerPoint",pdf:"PDF",xlsx:"Excel",txt:"文本",md:"Markdown",other:"综合"};
export const priorities = {normal:"普通",high:"高优先级",urgent:"紧急"};
export const roles = {admin:"超级管理员",manager:"部门主管",employee:"员工",reviewer:"审阅人"};
export const eventLabels = {
  "task.created":"创建任务", "task.published":"发布任务", "task.scheduled":"安排定时发布", "task.updated":"调整任务",
  "task.draft_updated":"修改任务草稿", "task.start":"开始办理", "task.pause":"暂停办理", "task.comment":"任务留言",
  "task.extension":"申请延期", "task.submitted":"提交成果", "task.approved":"验收通过",
  "task.returned":"退回修改", "task.revoke":"撤销任务", "task.close":"关闭任务", "file.uploaded":"上传文件",
  "file.material":"添加任务资料", "file.generated":"生成文件", "file.edited":"编辑文件", "file.copied":"关联成果",
  "file.preview":"预览附件", "file.download":"下载附件", "agent.started":"启动文稿协助",
  "agent.completed":"完成文稿协助", "agent.failed":"文稿协助失败", "schedule.created":"启动重复计划",
  "schedule.updated":"更新重复计划", "auth.login":"企业身份登录"
};

export function badge(value) {
  return `<span class="badge ${e(value)}"><i class="dot"></i>${e(statuses[value] || priorities[value] || value)}</span>`;
}

export function dateLabel(date, options = { month:"numeric", day:"numeric", weekday:"long" }) {
  if (!date) return "—";
  const simple = /^\d{4}-\d{2}-\d{2}$/.test(date);
  return new Intl.DateTimeFormat("zh-CN", { ...options, timeZone: simple ? "UTC" : (state.me?.timezone || state.config?.timezone || "Asia/Shanghai") })
    .format(new Date(simple ? date + "T12:00:00Z" : date));
}

export function deadline(value) {
  return value ? dateLabel(value, {month:"numeric",day:"numeric",hour:"2-digit",minute:"2-digit",hour12:false}).replace(" 24:", " 00:") : "无截止时间";
}

export function localInput(value) {
  if (!value) return "";
  const tz = state.me?.timezone || state.config?.timezone || "Asia/Shanghai";
  const pieces = new Intl.DateTimeFormat("sv-SE", {timeZone:tz,year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit",hourCycle:"h23"})
    .format(new Date(value));
  return pieces.replace(" ", "T");
}

export function fromLocal(value) {
  if (!value) return null;
  const tz = state.me?.timezone || state.config?.timezone || "Asia/Shanghai";
  const target = new Date(value + ":00Z").getTime();
  let ts = target;
  for (let round = 0; round < 3; round++) {
    const actual = new Date(localInput(new Date(ts).toISOString()) + ":00Z").getTime();
    const diff = target - actual;
    if (!diff) return new Date(ts).toISOString();
    ts += diff;
  }
  if (!tz) return new Date(ts).toISOString();
  throw new Error("所选当地时间不存在，请更改时间");
}

export function fileIcon(kind) {
  return `<span class="task-icon ${e(kind)}">${icon(kind)}</span>`;
}

export function button(label, action, opts = {}) {
  return `<button type="button" class="button ${opts.cls || ""}" data-act="${e(action)}" ${opts.id ? 'data-id="'+e(opts.id)+'"' : ""} ${opts.disabled ? "disabled" : ""}>${opts.icon ? icon(opts.icon) : ""}${e(label)}</button>`;
}

export function toastMessage(message) {
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => toast.classList.remove("show"), 3800);
}
export function modal(title, html) {
  dialog.innerHTML = `<header class="dialog-title"><h2>${e(title)}</h2><button class="icon-button" type="button" data-act="close-dialog" aria-label="关闭">${icon("x")}</button></header><div class="dialog-body">${html}</div>`;
  if (!dialog.open) dialog.showModal();
}
export function errorIn(node, error) {
  let target = node.querySelector(".error");
  if (!target) {
    target = document.createElement("p");
    target.className = "error";
    target.setAttribute("role", "alert");
    node.append(target);
  }
  target.textContent = error.message || String(error);
}

export async function api(path, options = {}) {
  const headers = new Headers(options.headers || {});
  let data = options.body;
  if (data !== undefined && !(data instanceof FormData)) {
    headers.set("Content-Type", "application/json");
    data = JSON.stringify(data);
  }
  const response = await fetch("/api/v1" + path, { ...options, headers, body: data, credentials:"same-origin" });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = Array.isArray(result.detail)
      ? result.detail.map((entry) => entry.msg || "请检查输入").slice(0, 3).join("、")
      : result.detail || `工作台请求失败（${response.status}）`;
    const error = new Error(message);
    error.status = response.status;
    throw error;
  }
  return result;
}

export function post(path, body, key = "") {
  return api(path, {method:"POST",body,headers:key ? {"Idempotency-Key":key} : {}});
}

export function go(path) {
  const hash = "#" + path;
  if (location.hash === hash) window.dispatchEvent(new Event("hashchange"));
  else location.hash = hash;
}

export function fileMarkup(file, opts = {}) {
  const size = file.size === undefined ? "" : file.size > 1048576 ? `${(file.size/1048576).toFixed(1)} MB` : `${Math.ceil(file.size/1024)} KB`;
  return `<div class="file">${opts.check ? `<input type="checkbox" name="${e(opts.check)}" value="${e(file.id)}" aria-label="${e(file.name)}" ${opts.checked ? "checked" : ""}>` : ""}${fileIcon(file.extension)}
    <span class="file-meta"><strong>${e(file.name)}</strong><small>${e(labels[file.extension] || file.extension)}${size ? " · " + size : ""}${file.parent_id ? " · 历史版本保留" : ""}</small></span>
    <button class="text-button" type="button" data-act="preview" data-id="${e(file.id)}">预览</button>
    <a class="text-button" href="/api/v1/files/${encodeURIComponent(file.id)}/download" download>下载</a>
    ${opts.edit && ["docx","pptx"].includes(file.extension) ? `<button class="text-button" type="button" data-act="edit-file" data-id="${e(file.id)}">修改</button>` : ""}
    ${opts.copy ? `<button class="text-button" type="button" data-act="copy-file" data-id="${e(file.id)}">关联</button>` : ""}
  </div>`;
}

export function uploadMarkup(summary = "拖入文件或点击添加") {
  return `<label class="button upload-zone" data-upload-zone>${icon("upload")} ${e(summary)} <input type="file" data-upload accept=".docx,.pptx,.xlsx,.pdf,.txt,.md" multiple></label>`;
}

export async function uploadSelected(files, assignment = "", target = "") {
  const uploaded = [];
  for (const file of files) {
    const form = new FormData();
    form.append("file", file);
    if (assignment) form.append("assignment_id", assignment);
    uploaded.push(await api(target || "/files", {method:"POST",body:form}));
  }
  return uploaded;
}

export function saveDraft(key, data) {
  try { sessionStorage.setItem("qo:" + state.me.id + ":" + key, JSON.stringify(data)); } catch {}
}

export function readDraft(key) {
  try { return JSON.parse(sessionStorage.getItem("qo:" + state.me.id + ":" + key) || "null"); } catch { return null; }
}

export function clearDrafts() {
  try {
    for (const key of Object.keys(sessionStorage)) if (key.startsWith("qo:")) sessionStorage.removeItem(key);
  } catch {}
}
