export const root = document.querySelector("#app");
export const dialog = document.querySelector("#dialog");
const toast = document.querySelector("#toast");

export const state = {
  config: null,
  me: null,
  engine: null,
  route: [],
  today: [],
  filter: "all",
  focused: null,
  agent: null,
  chatWorkspace: null,
  conversations: [],
  usageOpen: false,
  usageState: null,
  contextOpen: false,
  studioTab: "context",
  imageDrawerOpen: false,
  filesCache: [],
  filesFilter: "all",
  filesQuery: "",
  update: null,
  updateBannerDismissed: false,
  admin: { offset: 0, department: "", status: "", date: "" },
  auditOffset: 0,
  feedback: { scope: "auto", status: "", category: "", reporter_id: "", q: "", offset: 0 },
  feedbackOpenCount: 0
};

const paths = {
  mark: '<path class="logo-outer" d="M13 18H6V6h12v7" stroke-width="2"/><path class="logo-inner" d="M9.6 13.4V9.6h3.8" stroke-width="1.45"/><path class="logo-axis" d="M12.6 12.6L18.8 18.8" stroke-width="2"/>',
  message: '<path d="M5 4h14a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H9l-5 4V6a2 2 0 0 1 1-2Z"/><path d="M8 9h9m-9 4h6"/>',
  today: '<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M8 2v4m8-4v4M3 9h18m-13 6 2 2 5-5"/>',
  folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/>',
  history: '<path d="M3 12a9 9 0 1 1 3 6.7M3 16v-4h4m5-5v5l4 2"/>',
  sliders: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.6 1.6 0 0 1 .3 1.8l.1.1-1.6 1.6-.1-.1a1.6 1.6 0 0 1-1.8-.2 1.6 1.6 0 0 0-2.5 1v.2h-2.3v-.2a1.6 1.6 0 0 0-2.5-1 1.6 1.6 0 0 1-1.8.2l-.1.1L5.5 17l.1-.1a1.6 1.6 0 0 1 .2-1.8 1.6 1.6 0 0 0-1-2.5h-.2v-2.3h.2a1.6 1.6 0 0 0 1-2.5 1.6 1.6 0 0 1-.2-1.8l-.1-.1L7.1 4.3l.1.1A1.6 1.6 0 0 1 9 4.6a1.6 1.6 0 0 0 2.5-1v-.2h2.3v.2a1.6 1.6 0 0 0 2.5 1 1.6 1.6 0 0 1 1.8-.2l.1-.1 1.6 1.6-.1.1a1.6 1.6 0 0 1-.3 1.8 1.6 1.6 0 0 0 1 2.5h.2v2.3h-.2a1.6 1.6 0 0 0-1 2.4Z"/>',
  tasks: '<rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V2h6v2M8 10h8m-8 4h8m-8 4h5"/>',
  users: '<circle cx="9" cy="8" r="3"/><path d="M2 20v-2a4 4 0 0 1 4-4h6a4 4 0 0 1 4 4v2m1-15a3 3 0 0 1 0 6m5 9v-2a4 4 0 0 0-4-4"/>',
  audit: '<path d="M12 22s7-4 7-11V5l-7-3-7 3v6c0 7 7 11 7 11Z"/><path d="m9 11 2 2 4-4"/>',
  right: '<path d="m9 5 7 7-7 7"/>',
  down: '<path d="m5 9 7 7 7-7"/>',
  left: '<path d="m15 5-7 7 7 7"/>',
  arrow: '<path d="M5 12h14m-6-6 6 6-6 6"/>',
  search: '<circle cx="10.8" cy="10.8" r="6.8"/><path d="m16 16 4 4"/>',
  refresh: '<path d="M20 11a8 8 0 1 0-1.3 5m1.3-9v4h-4"/>',
  calendar: '<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M8 2v4m8-4v4M3 9h18"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
  checkPlain: '<path d="m5 12 5 5 9-10"/>',
  alert: '<path d="M10.3 4.5a2 2 0 0 1 3.4 0l7.5 13A2 2 0 0 1 19.4 20H4.6a2 2 0 0 1-1.8-2.5zM12 9v4m0 3h.01"/>',
  docx: '<path d="M6 3h8l4 4v14H6zM14 3v5h4m-9 4h6m-6 4h6"/>',
  pptx: '<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M12 17v4m-3 0h6M8 13v-3m4 3V8m4 5v-2"/>',
  xlsx: '<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M4 9h16M4 14h16M10 9v11"/>',
  upload: '<path d="M12 16V4m-5 5 5-5 5 5M4 15v5h16v-5"/>',
  download: '<path d="M12 4v12m-5-5 5 5 5-5M4 16v4h16v-4"/>',
  edit: '<path d="m4 16-1 5 5-1L20 8l-4-4ZM14 6l4 4"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  x: '<path d="m6 6 12 12M6 18 18 6"/>',
  logout: '<path d="M11 4H4v16h7m5-13 5 5-5 5m5-5H8"/>',
  link: '<path d="M13 5h6v6m0-6-8 8M9 5H5v14h14v-4"/>',
  clip: '<path d="m9 11 5-5a3 3 0 1 1 4 4l-8 8a4 4 0 1 1-6-6l8-8"/>',
  copy: '<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
  stop: '<rect x="6" y="6" width="12" height="12" rx="2"/>',
  play: '<polygon points="7 5 19 12 7 19 7 5"/>',
  sidebar: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M9 4v16"/>',
  panelRight: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M15 4v16"/>',
  menu: '<path d="M4 7h16M4 12h16M4 17h16"/>',
  sparkles: '<path d="m12 3 1.7 4.3L18 9l-4.3 1.7L12 15l-1.7-4.3L6 9l4.3-1.7L12 3Zm6 11 .9 2.1L21 17l-2.1.9L18 20l-.9-2.1L15 17l2.1-.9L18 14Z"/>',
  inbox: '<path d="M3 13V6a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v7l-3 7H6l-3-7Z"/><path d="M3 13h5l2 3h4l2-3h5"/>',
  layers: '<path d="m12 3 9 4.5-9 4.5-9-4.5L12 3Zm-9 9 9 4.5 9-4.5M3 16.5 12 21l9-4.5"/>',
  sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M4.9 4.9l1.4 1.4m11.4 11.4 1.4 1.4M2 12h2m16 0h2M4.9 19.1l1.4-1.4m11.4-11.4 1.4-1.4"/>',
  moon: '<path d="M20 14.5A8 8 0 0 1 9.5 4 8.5 8.5 0 1 0 20 14.5Z"/>',
  arrowUp: '<path d="M12 19V5m-6 6 6-6 6 6"/>',
  image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="8.5" cy="9.5" r="1.5"/><path d="m21 15-4.5-4.5L9 18"/>',
  slash: '<path d="M15 4 9 20"/>',
  maximize: '<path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/>',
  move: '<path d="M5 9l-3 3 3 3M9 5l3-3 3 3M15 19l-3 3-3-3M19 9l3 3-3 3M2 12h20M12 2v20"/>',
  crop: '<rect x="5" y="5" width="14" height="14" rx="1.5" stroke-dasharray="3 2.5"/>',
  bug: '<path d="M8 7a4 4 0 0 1 8 0v1H8V7Z"/><rect x="6" y="8" width="12" height="11" rx="5"/><path d="M12 8v11M3 11h3m12 0h3M4 16l2.5-1M20 16l-2.5-1M9 4 7.5 2.5M15 4l1.5-1.5"/>',
  flag: '<path d="M5 21V4h11l-1.5 4.5L16 13H5"/>'
};

const THEME_KEY = "qo-theme";

export function getTheme() {
  try {
    const saved = localStorage.getItem(THEME_KEY);
    if (saved === "light" || saved === "dark") return saved;
  } catch {}
  return "light";
}

export function applyTheme(mode = getTheme()) {
  const next = mode === "dark" ? "dark" : "light";
  document.documentElement.dataset.theme = next;
  if (window.officeDesktop?.hasTitleBarOverlay) {
    document.documentElement.classList.add("desktop-titlebar");
  }
  try { localStorage.setItem(THEME_KEY, next); } catch {}
  try { void window.officeDesktop?.setTheme?.(next); } catch {}
  return next;
}

export function toggleTheme() {
  return applyTheme(getTheme() === "dark" ? "light" : "dark");
}

export function icon(name) {
  const isMark = name === "mark";
  return `<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${isMark ? "2" : "1.5"}" stroke-linecap="${isMark ? "butt" : "round"}" stroke-linejoin="${isMark ? "miter" : "round"}">${paths[name] || paths.docx}</svg>`;
}

export function e(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
}

export function formatBytes(bytes) {
  const value = Number(bytes);
  if (!Number.isFinite(value) || value <= 0) return "0 B";
  if (value < 1024) return `${Math.round(value)} B`;
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(value < 10 * 1024 ? 1 : 0)} KB`;
  if (value < 1024 * 1024 * 1024) return `${(value / (1024 * 1024)).toFixed(1)} MB`;
  return `${(value / (1024 * 1024 * 1024)).toFixed(2)} GB`;
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
  overdue:"已逾期",draft:"草稿",active:"运行中",finished:"已结束",
  open:"待处理",resolved:"已解决"
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
  "schedule.updated":"更新重复计划", "auth.login":"企业身份登录", "image.started":"启动图片生成",
  "image.generated":"生成图片", "image.ai_generated":"AI 生成图片", "file.image":"查看图片",
  "feedback.submitted":"提交问题反馈", "feedback.updated":"更新反馈状态"
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
  return `<span class="task-icon ${e(kind || "other")}">${icon(kind)}</span>`;
}

export function button(label, action, opts = {}) {
  return `<button type="button" class="button ${opts.cls || ""}" data-act="${e(action)}" ${opts.id ? 'data-id="'+e(opts.id)+'"' : ""} ${opts.title ? 'title="'+e(opts.title)+'"' : ""} ${opts.disabled ? "disabled" : ""}>${opts.icon ? icon(opts.icon) : ""}${e(label)}</button>`;
}

export function emptyState(title, copy = "", glyph = "inbox", action = "") {
  return `<div class="empty-state">${icon(glyph)}<strong>${e(title)}</strong>${copy ? `<p>${e(copy)}</p>` : ""}${action}</div>`;
}

export function toastMessage(message, tone = "info") {
  toast.textContent = message;
  toast.classList.remove("toast-error", "toast-success");
  if (tone === "error") toast.classList.add("toast-error");
  if (tone === "success") toast.classList.add("toast-success");
  toast.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => toast.classList.remove("show"), 3600);
}

export async function copyText(text, okMessage = "已复制到剪贴板") {
  if (!text) return false;
  // 桌面端会话拒绝全部权限检查，clipboard.writeText 可能被拒；被拒时改用 execCommand，并以其返回值判定是否成功。
  let copied = false;
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      copied = true;
    }
  } catch {}
  if (!copied) {
    const area = document.createElement("textarea");
    try {
      area.value = text;
      area.className = "sr-only";
      document.body.append(area);
      area.select();
      copied = document.execCommand("copy") === true;
    } catch {
      copied = false;
    } finally {
      area.remove();
    }
  }
  if (copied) toastMessage(okMessage, "success");
  else toastMessage("复制失败，请手动选择文本复制", "error");
  return copied;
}

export function modal(title, html) {
  dialog.innerHTML = `<header class="dialog-title"><h2>${e(title)}</h2><button class="icon-button" type="button" data-act="close-dialog" aria-label="关闭">${icon("x")}</button></header><div class="dialog-body">${html}</div>`;
  if (!dialog.open) dialog.showModal();
}

export function cleanErrorMessage(error, fallback = "操作失败") {
  const raw = typeof error === "string" ? error : (error?.message || String(error || ""));
  const cleaned = raw.replace(/^Error invoking remote method '[^']+':\s*(?:Error:\s*)?/i, "").trim();
  return cleaned || fallback;
}

export function errorIn(node, error) {
  let target = node.querySelector(".error");
  if (!target) {
    target = document.createElement("p");
    target.className = "error";
    target.setAttribute("role", "alert");
    node.append(target);
  }
  target.textContent = cleanErrorMessage(error);
}

let apiErrorHook = null;
export function setApiErrorHook(fn) {
  apiErrorHook = typeof fn === "function" ? fn : null;
}

export async function api(path, options = {}) {
  const headers = new Headers(options.headers || {});
  let data = options.body;
  if (data !== undefined && !(data instanceof FormData)) {
    headers.set("Content-Type", "application/json");
    data = JSON.stringify(data);
  }
  const method = (options.method || "GET").toUpperCase();
  const startedAt = Date.now();
  let response;
  try {
    response = await fetch("/api/v1" + path, { ...options, headers, body: data, credentials:"same-origin" });
  } catch (netErr) {
    try { apiErrorHook?.({ method, path, status: 0, durationMs: Date.now() - startedAt, message: netErr?.message || "网络请求异常" }); } catch {}
    throw netErr;
  }
  const result = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = Array.isArray(result.detail)
      ? result.detail.map((entry) => entry.msg || "请检查输入").slice(0, 3).join("、")
      : result.detail || `工作台请求失败（${response.status}）`;
    const error = new Error(message);
    error.status = response.status;
    if (response.status !== 401 && !String(path).startsWith("/feedback")) {
      try { apiErrorHook?.({ method, path, status: response.status, durationMs: Date.now() - startedAt, message }); } catch {}
    }
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
    <span class="file-actions">
      <button class="text-button" type="button" data-act="preview" data-id="${e(file.id)}">预览</button>
      <a class="text-button" href="/api/v1/files/${encodeURIComponent(file.id)}/download" download>下载</a>
      ${opts.edit && ["docx","pptx"].includes(file.extension) ? `<button class="text-button" type="button" data-act="edit-file" data-id="${e(file.id)}">修改</button>` : ""}
      ${opts.copy ? `<button class="text-button" type="button" data-act="copy-file" data-id="${e(file.id)}">关联</button>` : ""}
    </span>
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

/* 全站自定义下拉菜单引擎：自动接管所有原生 <select>，保持 FormData / value / change / disabled 完全兼容 */
let customSelectInitialized = false;
let activeSelectMenu = null;
let activeSelectOwner = null;
const selectTriggerMap = new WeakMap();
const selectObserverMap = new WeakMap();

function formatSelectOptionHtml(rawText) {
  const text = String(rawText ?? "");
  const tagged = /^\[(图片|文字|图形)\]\s*(.*)$/.exec(text);
  if (!tagged) return `<span class="qo-select-text">${e(text)}</span>`;
  const kindMap = { 图片: "image", 文字: "text", 图形: "shape" };
  return `<span class="qo-opt-tag is-${kindMap[tagged[1]] || "shape"}">${e(tagged[1])}</span><span class="qo-select-text">${e(tagged[2])}</span>`;
}

export function closeCustomSelectMenu({ restoreFocus = false } = {}) {
  const prev = activeSelectOwner;
  if (activeSelectMenu) {
    activeSelectMenu.remove();
    activeSelectMenu = null;
  }
  if (prev) {
    const trigger = selectTriggerMap.get(prev);
    if (trigger) {
      trigger.setAttribute("aria-expanded", "false");
      trigger.closest(".qo-select-wrap")?.classList.remove("is-open");
      if (restoreFocus && trigger.isConnected) trigger.focus({ preventScroll: true });
    }
    activeSelectOwner = null;
  }
}

function syncSelectTrigger(select) {
  if (!select) return;
  const trigger = selectTriggerMap.get(select);
  if (!trigger) return;
  const wrap = trigger.closest(".qo-select-wrap");
  const selectedOpt = select.options?.[select.selectedIndex] || [...(select.options || [])].find((o) => o.value === select.value) || select.options?.[0];
  const labelText = selectedOpt ? selectedOpt.textContent : "";
  const labelNode = trigger.querySelector(".qo-select-label");
  if (labelNode) labelNode.innerHTML = formatSelectOptionHtml(labelText);
  trigger.disabled = Boolean(select.disabled);
  trigger.title = labelText || "";
  if (select.getAttribute("aria-label")) {
    trigger.setAttribute("aria-label", select.getAttribute("aria-label"));
  }
  if (wrap) {
    wrap.classList.toggle("is-disabled", Boolean(select.disabled));
    wrap.hidden = Boolean(select.hidden);
  }
  if (activeSelectOwner === select && select.disabled) {
    closeCustomSelectMenu();
  }
}

function positionSelectMenu(menu, trigger) {
  const rect = trigger.getBoundingClientRect();
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  const margin = 10;
  const minW = Math.max(rect.width, 168);
  const maxW = Math.min(Math.max(rect.width + 64, 340), vw - margin * 2);
  menu.style.minWidth = `${Math.round(minW)}px`;
  menu.style.maxWidth = `${Math.round(maxW)}px`;

  const spaceBelow = vh - rect.bottom - margin;
  const spaceAbove = rect.top - margin;
  const openUp = spaceBelow < 220 && spaceAbove > spaceBelow;
  const maxH = Math.max(120, Math.min(320, (openUp ? spaceAbove : spaceBelow) - 8));
  menu.style.maxHeight = `${Math.round(maxH)}px`;
  menu.classList.toggle("open-up", openUp);

  const menuRect = menu.getBoundingClientRect();
  let left = rect.left;
  if (left + menuRect.width > vw - margin) {
    left = Math.max(margin, rect.right - menuRect.width);
  }
  if (left < margin) left = margin;
  menu.style.left = `${Math.round(left)}px`;

  if (openUp) {
    menu.style.top = "auto";
    menu.style.bottom = `${Math.round(vh - rect.top + 6)}px`;
  } else {
    menu.style.bottom = "auto";
    menu.style.top = `${Math.round(rect.bottom + 6)}px`;
  }
}

function openCustomSelectMenu(select, trigger) {
  if (!select || select.disabled || !select.options.length) return;
  if (activeSelectOwner === select && activeSelectMenu) {
    closeCustomSelectMenu();
    return;
  }
  closeCustomSelectMenu();

  const menu = document.createElement("div");
  menu.className = "qo-select-menu";
  menu.setAttribute("role", "listbox");
  if (select.getAttribute("aria-label")) {
    menu.setAttribute("aria-label", select.getAttribute("aria-label"));
  }

  const opts = [...select.options];
  const currentIdx = select.selectedIndex;
  for (let i = 0; i < opts.length; i += 1) {
    const opt = opts[i];
    const isSelected = i === currentIdx;
    const item = document.createElement("button");
    item.type = "button";
    item.className = `qo-select-option${isSelected ? " active" : ""}`;
    item.setAttribute("role", "option");
    item.setAttribute("aria-selected", isSelected ? "true" : "false");
    item.dataset.index = String(i);
    if (opt.disabled) item.disabled = true;
    item.innerHTML = `<span class="qo-select-option-main">${formatSelectOptionHtml(opt.textContent)}</span><span class="qo-select-check">${isSelected ? icon("checkPlain") : ""}</span>`;
    item.addEventListener("click", (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      if (opt.disabled) return;
      const changed = select.selectedIndex !== i || select.value !== opt.value;
      select.selectedIndex = i;
      syncSelectTrigger(select);
      closeCustomSelectMenu({ restoreFocus: true });
      if (changed) {
        select.dispatchEvent(new Event("input", { bubbles: true }));
        select.dispatchEvent(new Event("change", { bubbles: true }));
      }
    });
    menu.append(item);
  }

  const hostDialog = select.closest("dialog[open]");
  (hostDialog || document.body).append(menu);
  activeSelectMenu = menu;
  activeSelectOwner = select;
  trigger.setAttribute("aria-expanded", "true");
  trigger.closest(".qo-select-wrap")?.classList.add("is-open");
  positionSelectMenu(menu, trigger);

  const activeItem = menu.querySelector(".qo-select-option.active");
  if (activeItem) activeItem.scrollIntoView({ block: "nearest" });
}

function enhanceSelect(select) {
  if (!select || select.tagName !== "SELECT" || select.multiple || Number(select.size) > 1) return;
  if (select.dataset.qoSelectReady === "1") {
    syncSelectTrigger(select);
    return;
  }
  const parent = select.parentNode;
  if (!parent) return;

  select.dataset.qoSelectReady = "1";
  const wrap = document.createElement("span");
  wrap.className = "qo-select-wrap";
  for (const cls of ["deck-select", "deck-page-select", "deck-shape-select", "form-control"]) {
    if (select.classList.contains(cls)) wrap.classList.add(cls);
  }
  if (select.closest(".image-controls")) wrap.classList.add("is-compact");

  const trigger = document.createElement("button");
  trigger.type = "button";
  trigger.className = "qo-select-trigger";
  trigger.setAttribute("aria-haspopup", "listbox");
  trigger.setAttribute("aria-expanded", "false");
  trigger.innerHTML = `<span class="qo-select-label"></span><span class="qo-select-caret">${icon("down")}</span>`;

  parent.insertBefore(wrap, select);
  wrap.append(trigger, select);
  select.classList.add("qo-select-native");
  select.tabIndex = -1;
  select.setAttribute("aria-hidden", "true");
  selectTriggerMap.set(select, trigger);

  trigger.addEventListener("click", (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    openCustomSelectMenu(select, trigger);
  });
  trigger.addEventListener("keydown", (ev) => {
    if (ev.key === "ArrowDown" || ev.key === "ArrowUp" || ev.key === "Enter" || ev.key === " ") {
      if (activeSelectOwner !== select) {
        ev.preventDefault();
        openCustomSelectMenu(select, trigger);
      }
    }
  });
  select.addEventListener("change", () => syncSelectTrigger(select));

  const obs = new MutationObserver(() => {
    syncSelectTrigger(select);
    if (activeSelectOwner === select) closeCustomSelectMenu();
  });
  obs.observe(select, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ["disabled", "hidden", "aria-label"] });
  selectObserverMap.set(select, obs);
  syncSelectTrigger(select);
}

export function enhanceSelectsIn(container = document) {
  if (!container?.querySelectorAll) return;
  for (const select of container.querySelectorAll("select:not([data-qo-select-ready='1'])")) {
    enhanceSelect(select);
  }
}

export function initCustomSelects() {
  if (customSelectInitialized || typeof document === "undefined") return;
  customSelectInitialized = true;

  // 监听通过 JS 直接赋值 select.value / select.selectedIndex，自动同步自定义触发器文案
  for (const prop of ["value", "selectedIndex"]) {
    const desc = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, prop);
    if (desc?.set && desc?.get) {
      Object.defineProperty(HTMLSelectElement.prototype, prop, {
        configurable: true,
        enumerable: desc.enumerable,
        get() { return desc.get.call(this); },
        set(val) {
          desc.set.call(this, val);
          if (this.dataset?.qoSelectReady === "1") syncSelectTrigger(this);
        }
      });
    }
  }

  enhanceSelectsIn(document);
  const rootObserver = new MutationObserver((mutations) => {
    for (const m of mutations) {
      for (const node of m.addedNodes) {
        if (node.nodeType !== 1) continue;
        if (node.tagName === "SELECT") enhanceSelect(node);
        else if (node.querySelectorAll) enhanceSelectsIn(node);
      }
    }
  });
  rootObserver.observe(document.body, { childList: true, subtree: true });

  document.addEventListener("pointerdown", (ev) => {
    if (!activeSelectMenu) return;
    if (ev.target.closest?.(".qo-select-menu") || ev.target.closest?.(".qo-select-trigger")) return;
    closeCustomSelectMenu();
  }, true);

  document.addEventListener("keydown", (ev) => {
    if (!activeSelectMenu || !activeSelectOwner) return;
    const options = [...activeSelectMenu.querySelectorAll(".qo-select-option:not(:disabled)")];
    if (ev.key === "Escape") {
      ev.preventDefault();
      ev.stopPropagation();
      closeCustomSelectMenu({ restoreFocus: true });
      return;
    }
    if (!options.length) return;
    const cur = options.findIndex((el) => el.classList.contains("is-focused") || el.classList.contains("active"));
    if (ev.key === "ArrowDown" || ev.key === "ArrowUp") {
      ev.preventDefault();
      const next = (cur + (ev.key === "ArrowDown" ? 1 : -1) + options.length) % options.length;
      options.forEach((el, idx) => el.classList.toggle("is-focused", idx === next));
      options[next].scrollIntoView({ block: "nearest" });
    } else if (ev.key === "Enter" || ev.key === " ") {
      ev.preventDefault();
      const target = options[cur >= 0 ? cur : 0];
      target?.click();
    }
  });

  window.addEventListener("resize", () => {
    if (activeSelectMenu && activeSelectOwner) {
      const trigger = selectTriggerMap.get(activeSelectOwner);
      if (trigger?.isConnected) positionSelectMenu(activeSelectMenu, trigger);
      else closeCustomSelectMenu();
    }
  });
  window.addEventListener("scroll", (ev) => {
    if (!activeSelectMenu) return;
    if (ev.target === activeSelectMenu || activeSelectMenu.contains(ev.target)) return;
    closeCustomSelectMenu();
  }, true);
}
