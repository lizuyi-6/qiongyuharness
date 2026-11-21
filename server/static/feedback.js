import { api, badge, button, cleanErrorMessage, copyText, deadline, dialog, e, emptyState,
  formatBytes, icon, isManager, modal, post, roles, setApiErrorHook, state, toastMessage } from "./ui.js";
import { detectDesktopVersion, pageHead, pagination } from "./views.js";

const MAX_CLIENT_LOGS = 250;
const PAGE_SIZE = 20;
const clientLogs = [];
let captureInitialized = false;
let lastErrorContext = null;
let lastLoadedFeedback = null;

export const feedbackCategories = {
  bug: "功能异常",
  ai: "AI 执行/生成失败",
  ppt: "PPT 排版与画布",
  update: "安装与更新",
  suggestion: "体验建议",
  other: "其他问题"
};

export const feedbackStatuses = {
  open: "待处理",
  in_progress: "处理中",
  resolved: "已解决",
  closed: "已关闭"
};

function redactClientText(value, maxLen = 1500) {
  let text = String(value ?? "");
  if (!text) return "";
  text = text.replace(/\b(Bearer\s+)[A-Za-z0-9._~+/=-]{8,}/gi, "$1[REDACTED]");
  text = text.replace(/\b(sk-[A-Za-z0-9_-]{8,})/g, "[REDACTED_KEY]");
  text = text.replace(/\b(qo_session|qo_browser)=([^;\s"'`]+)/gi, "$1=[REDACTED]");
  text = text.replace(/data:image\/[a-zA-Z0-9.+-]+;base64,[A-Za-z0-9+/=]{64,}/g, "[IMAGE_BASE64_OMITTED]");
  return text.length > maxLen ? text.slice(0, maxLen) + "…" : text;
}

export function recordClientLog(level, scope, message, detail = null) {
  const safeMsg = redactClientText(message, 1200);
  if (!safeMsg) return;
  const entry = {
    ts: new Date().toISOString(),
    level: String(level || "info").toLowerCase(),
    scope: String(scope || "ui").slice(0, 32),
    message: safeMsg,
    route: location.hash || "#"
  };
  if (detail && typeof detail === "object") {
    try {
      const cleaned = {};
      for (const [k, v] of Object.entries(detail).slice(0, 12)) {
        if (/(token|key|secret|password|cookie|authorization|base64)/i.test(k)) cleaned[k] = "[REDACTED]";
        else if (typeof v === "string") cleaned[k] = redactClientText(v, 400);
        else if (typeof v === "number" || typeof v === "boolean" || v === null) cleaned[k] = v;
      }
      if (Object.keys(cleaned).length) entry.detail = cleaned;
    } catch {}
  } else if (typeof detail === "string" && detail) {
    entry.detail = { text: redactClientText(detail, 400) };
  }
  clientLogs.push(entry);
  if (clientLogs.length > MAX_CLIENT_LOGS) clientLogs.splice(0, clientLogs.length - MAX_CLIENT_LOGS);
  if (entry.level === "error") {
    lastErrorContext = {
      ts: entry.ts,
      scope: entry.scope,
      message: safeMsg,
      route: entry.route,
      conversationId: state.chatWorkspace?.session?.id || ""
    };
  }
}

export function getRecentClientLogs() {
  return clientLogs.slice(-MAX_CLIENT_LOGS);
}

export function getLastErrorContext() {
  return lastErrorContext ? { ...lastErrorContext } : null;
}

export function initClientLogCapture() {
  if (captureInitialized) return;
  captureInitialized = true;

  setApiErrorHook((info) => {
    recordClientLog("error", "api", `${info.method} ${info.path} -> HTTP ${info.status}: ${info.message}`, {
      status: info.status,
      durationMs: info.durationMs
    });
  });

  window.addEventListener("error", (event) => {
    const msg = event.message || event.error?.message || "未捕获前端脚本错误";
    recordClientLog("error", "window.error", msg, {
      file: event.filename || "",
      line: event.lineno || 0,
      col: event.colno || 0,
      stack: event.error?.stack || ""
    });
  });

  window.addEventListener("unhandledrejection", (event) => {
    const reason = event.reason;
    const msg = cleanErrorMessage(reason, "未处理的异步异常");
    recordClientLog("error", "promise.rejection", msg, {
      stack: reason?.stack || ""
    });
  });

  const origError = console.error?.bind(console);
  if (origError) {
    console.error = (...args) => {
      try {
        const text = args.map((a) => (a instanceof Error ? `${a.name}: ${a.message}` : typeof a === "object" ? JSON.stringify(a) : String(a))).join(" ");
        recordClientLog("error", "console", text);
      } catch {}
      return origError(...args);
    };
  }
  const origWarn = console.warn?.bind(console);
  if (origWarn) {
    console.warn = (...args) => {
      try {
        const text = args.map((a) => (typeof a === "object" ? JSON.stringify(a) : String(a))).join(" ");
        recordClientLog("warn", "console", text);
      } catch {}
      return origWarn(...args);
    };
  }
  recordClientLog("info", "app", "client log capture initialized", {
    desktop: typeof window.officeDesktop?.runAgent === "function",
    version: detectDesktopVersion() || ""
  });
}

export function clearFeedbackCache() {
  lastLoadedFeedback = null;
}

function activeConversationId(preset = {}) {
  if (preset.conversationId) return String(preset.conversationId);
  if (state.route[0] === "chat") return String(state.chatWorkspace?.session?.id || "");
  return "";
}

async function gatherDiagnosticsPayload(preset = {}) {
  const conversationId = activeConversationId(preset);
  let desktopDiagnostics = null;
  if (typeof window.officeDesktop?.collectLogs === "function") {
    try {
      desktopDiagnostics = await window.officeDesktop.collectLogs({
        conversationId,
        userId: state.me?.id || ""
      });
    } catch (err) {
      recordClientLog("warn", "feedback", "collectLogs IPC failed: " + cleanErrorMessage(err));
    }
  }
  const clientVersion = detectDesktopVersion() || desktopDiagnostics?.app_version || "";
  const platform = desktopDiagnostics?.os_release || desktopDiagnostics?.platform || navigator.userAgent || "";
  const extraContext = {
    online: navigator.onLine,
    viewport: `${window.innerWidth}x${window.innerHeight}`,
    theme: document.documentElement.dataset.theme || "light",
    engine: state.engine?.engine || "",
    active_route: state.route.join("/"),
    ...(preset.extra && typeof preset.extra === "object" ? preset.extra : {})
  };
  return {
    conversationId,
    clientVersion,
    platform,
    desktopDiagnostics,
    clientLogs: getRecentClientLogs(),
    extraContext
  };
}

function guessCategory(preset = {}) {
  if (preset.category && feedbackCategories[preset.category]) return preset.category;
  const err = String(preset.errorMessage || "").toLowerCase();
  if (err.includes("更新") || err.includes("update") || err.includes("patch")) return "update";
  if (err.includes("ppt") || err.includes("画布") || err.includes("幻灯片")) return "ppt";
  if (preset.source === "chat" || preset.source === "checkpoint" || err.includes("模型") || err.includes("超时") || err.includes("agent")) return "ai";
  return "bug";
}

export async function quickUploadFeedback(preset = {}, triggerButton = null) {
  if (!state.me) throw new Error("请先登录工作台后再上传日志");
  const prevHtml = triggerButton ? triggerButton.innerHTML : "";
  if (triggerButton) {
    triggerButton.disabled = true;
    triggerButton.innerHTML = `${icon("refresh")}<span>正在打包上传日志…</span>`;
  }
  try {
    const diag = await gatherDiagnosticsPayload(preset);
    const category = guessCategory(preset);
    const errorMessage = cleanErrorMessage(preset.errorMessage || lastErrorContext?.message || "", "");
    const title = (preset.title || "").trim()
      || (errorMessage ? `一键日志上报：${errorMessage.slice(0, 64)}` : `${feedbackCategories[category]}（一键日志上报）`);
    const created = await post("/feedback", {
      category,
      title,
      description: (preset.description || "").trim(),
      error_message: errorMessage,
      conversation_id: diag.conversationId,
      route: location.hash || ("#" + state.route.join("/")),
      client_version: diag.clientVersion,
      platform: diag.platform,
      client_logs: diag.clientLogs,
      desktop_diagnostics: diag.desktopDiagnostics,
      extra_context: diag.extraContext
    });
    recordClientLog("info", "feedback", `quick uploaded report ${created.id}`);
    toastMessage(`日志已上传（编号 #${created.id.slice(0, 8)} · 记录对应人：${created.reporter_name}）`, "success");
    if (triggerButton) {
      triggerButton.innerHTML = `${icon("checkPlain")}<span>已上传日志</span>`;
      triggerButton.dataset.act = "feedback-view";
      triggerButton.dataset.id = created.id;
      triggerButton.disabled = false;
    }
    return created;
  } catch (error) {
    if (triggerButton) {
      triggerButton.innerHTML = prevHtml;
      triggerButton.disabled = false;
    }
    throw error;
  }
}

export function openFeedbackModal(preset = {}) {
  const me = state.me;
  if (!me) {
    toastMessage("请先登录后再提交问题反馈", "error");
    return;
  }
  const category = guessCategory(preset);
  const errorMessage = cleanErrorMessage(preset.errorMessage || (preset.useLastError ? lastErrorContext?.message : "") || "", "");
  const conversationId = activeConversationId(preset);
  const defaultTitle = (preset.title || "").trim() || (errorMessage ? errorMessage.slice(0, 72) : "");
  const isDesktop = typeof window.officeDesktop?.collectLogs === "function";
  const logHintParts = [
    isDesktop ? "桌面端运行日志" : "浏览器环境信息",
    `前端日志 ${clientLogs.length} 条`,
    conversationId ? "当前会话运行记录" : ""
  ].filter(Boolean);

  modal("问题反馈与一键上传日志", `<form class="stack feedback-form" data-form="feedback-submit" data-conversation-id="${e(conversationId)}">
    <div class="feedback-reporter-card">
      <span class="avatar">${e(me.name.slice(-2))}</span>
      <div class="feedback-reporter-meta">
        <strong>对应人：${e(me.name)}</strong>
        <small>${e(me.department || "未分配部门")} · ${e(roles[me.role] || me.role)}${me.email ? ` · ${e(me.email)}` : ""}</small>
      </div>
      <span class="feedback-auto-tag">${icon("checkPlain")}自动关联身份</span>
    </div>
    <div class="form-grid">
      <label class="field">问题分类
        <select name="category">
          ${Object.entries(feedbackCategories).map(([k, v]) => `<option value="${e(k)}" ${k === category ? "selected" : ""}>${e(v)}</option>`).join("")}
        </select>
      </label>
      <label class="field">简要标题（选填，留空自动生成）
        <input name="title" maxlength="160" value="${e(defaultTitle)}" placeholder="例如：生成 PPT 第 5 页排版错乱 / AI 响应超时">
      </label>
      ${errorMessage ? `<label class="field span">捕获到的错误信息
        <textarea name="error_message" rows="2" maxlength="2000" class="mono">${e(errorMessage)}</textarea>
      </label>` : `<input type="hidden" name="error_message" value="">`}
      <label class="field span">问题现象或复现步骤（选填，不填也可直接一键上传日志）
        <textarea name="description" rows="3" maxlength="6000" placeholder="可简述当时正在做什么、期望的结果与实际现象…">${e(preset.description || "")}</textarea>
      </label>
    </div>
    <div class="feedback-log-banner">
      <div class="feedback-log-banner-main">
        <strong>${icon("upload")}将自动打包上传诊断日志</strong>
        <span>${e(logHintParts.join(" · "))}（已自动脱敏，不含 API 密钥、密码与登录凭证）</span>
      </div>
      <label class="feedback-log-check">
        <input type="checkbox" name="include_logs" value="1" checked>
        <span>附带运行日志</span>
      </label>
    </div>
    <div class="actions between">
      <button type="button" class="text-button" data-act="route" data-id="feedback">查看历史反馈</button>
      <div class="actions">
        <button type="button" class="button" data-act="close-dialog">取消</button>
        <button type="submit" class="button primary">${icon("upload")}一键上传日志并提交</button>
      </div>
    </div>
  </form>`);
}

export async function submitFeedbackForm(form) {
  const data = new FormData(form);
  const includeLogs = data.get("include_logs") === "1";
  const conversationId = form.dataset.conversationId || "";
  const diag = includeLogs
    ? await gatherDiagnosticsPayload({ conversationId })
    : {
        conversationId,
        clientVersion: detectDesktopVersion() || "",
        platform: navigator.userAgent || "",
        desktopDiagnostics: null,
        clientLogs: [],
        extraContext: { logs_omitted_by_user: true }
      };
  const created = await post("/feedback", {
    category: String(data.get("category") || "bug"),
    title: String(data.get("title") || "").trim(),
    description: String(data.get("description") || "").trim(),
    error_message: String(data.get("error_message") || "").trim(),
    conversation_id: diag.conversationId,
    route: location.hash || ("#" + state.route.join("/")),
    client_version: diag.clientVersion,
    platform: diag.platform,
    client_logs: diag.clientLogs,
    desktop_diagnostics: diag.desktopDiagnostics,
    extra_context: diag.extraContext
  });
  dialog.close();
  recordClientLog("info", "feedback", `submitted feedback ${created.id}`);
  toastMessage(`反馈已提交（#${created.id.slice(0, 8)} · 已记录对应人：${created.reporter_name}）`, "success");
  return created;
}

export async function feedbackPage() {
  const f = state.feedback;
  const params = new URLSearchParams({
    scope: f.scope || "auto",
    offset: String(f.offset || 0),
    limit: String(PAGE_SIZE)
  });
  if (f.status) params.set("status", f.status);
  if (f.category) params.set("category", f.category);
  if (f.reporter_id) params.set("reporter_id", f.reporter_id);
  if (f.q) params.set("q", f.q);

  const data = await api("/feedback?" + params.toString());
  lastLoadedFeedback = data;
  state.feedbackOpenCount = Number(data.counts?.open || 0);
  const counts = data.counts || { all: 0, open: 0, in_progress: 0, resolved: 0, closed: 0 };
  const canManage = Boolean(data.can_manage);
  const scope = data.scope || (canManage ? "all" : "mine");
  const items = Array.isArray(data.items) ? data.items : [];
  const reporters = Array.isArray(data.reporters) ? data.reporters : [];

  const statusTabs = [
    ["", "全部", counts.all],
    ["open", "待处理", counts.open],
    ["in_progress", "处理中", counts.in_progress],
    ["resolved", "已解决", counts.resolved],
    ["closed", "已关闭", counts.closed]
  ];

  return `${pageHead(
    "问题反馈与日志",
    canManage ? "查看全员上报的问题、定位发生对应人并下载脱敏诊断日志，或一键上报当前日志。" : "遇到异常可一键打包上传桌面端与前端日志，自动记录对应人便于快速定位排查。",
    `<div class="actions">
      ${button("一键上传当前日志", "feedback-quick-upload", { icon: "upload" })}
      ${button("填写问题反馈", "feedback-open", { cls: "primary", icon: "bug" })}
    </div>`
  )}
  <section class="stack feedback-page">
    ${canManage ? `<div class="feedback-scope-bar">
      <div class="tabs" role="tablist" aria-label="反馈范围">
        <button type="button" class="tab ${scope === "all" ? "active" : ""}" data-act="feedback-scope" data-id="all">全部反馈</button>
        <button type="button" class="tab ${scope === "mine" ? "active" : ""}" data-act="feedback-scope" data-id="mine">我提交的</button>
      </div>
    </div>` : ""}
    <div class="tabs" aria-label="反馈状态">
      ${statusTabs.map(([key, label, cnt]) => `
        <button type="button" class="tab ${(f.status || "") === key ? "active" : ""}" data-act="feedback-status" data-id="${e(key)}">
          ${e(label)}<span class="tab-count">${cnt || 0}</span>
        </button>`).join("")}
    </div>
    <form class="filter" data-form="feedback-filter" aria-label="筛选问题反馈">
      ${canManage && scope === "all" && reporters.length ? `<select name="reporter_id" aria-label="对应人">
        <option value="">全部对应人（${reporters.length} 人）</option>
        ${reporters.map((r) => `<option value="${e(r.id)}" ${f.reporter_id === r.id ? "selected" : ""}>${e(r.name)}${r.department ? ` · ${e(r.department)}` : ""}（${r.count}）</option>`).join("")}
      </select>` : ""}
      <select name="category" aria-label="问题分类">
        <option value="">全部分类</option>
        ${Object.entries(feedbackCategories).map(([k, v]) => `<option value="${e(k)}" ${f.category === k ? "selected" : ""}>${e(v)}</option>`).join("")}
      </select>
      <input type="search" name="q" value="${e(f.q || "")}" placeholder="${canManage ? "搜索标题、报错或对应人姓名/邮箱…" : "搜索标题或错误信息…"}">
      <button class="button sm" type="submit">${icon("search")}筛选</button>
      ${(f.category || f.reporter_id || f.q || f.status) ? button("清空筛选", "feedback-reset", { cls: "sm" }) : ""}
    </form>
    ${items.length ? `<div class="table-wrap"><table class="feedback-list">
      <thead>
        <tr>
          <th>问题与报错摘要</th>
          <th>对应人</th>
          <th>环境与日志</th>
          <th>提交时间</th>
          <th>状态</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody>
        ${items.map((item) => `<tr>
          <td>
            <div class="feedback-cell-main">
              <div class="feedback-cell-title">
                <span class="feedback-cat-tag">${e(item.category_label || item.category)}</span>
                <button type="button" class="text-button feedback-title-btn" data-act="feedback-view" data-id="${e(item.id)}"><strong>${e(item.title)}</strong></button>
              </div>
              ${item.error_message ? `<p class="feedback-err-preview mono">${e(item.error_message.slice(0, 140))}</p>` : (item.description ? `<p class="feedback-desc-preview">${e(item.description.slice(0, 120))}</p>` : "")}
              ${item.conversation_title ? `<small class="hint">关联会话：${e(item.conversation_title)}</small>` : ""}
            </div>
          </td>
          <td>
            <div class="feedback-person">
              <span class="avatar sm">${e((item.reporter_name || "?").slice(-2))}</span>
              <div>
                <strong>${e(item.reporter_name)}</strong>
                <small>${e(item.reporter_department || "未分配")}${item.reporter_email ? ` · ${e(item.reporter_email)}` : ""}</small>
              </div>
            </div>
          </td>
          <td>
            <div class="feedback-env">
              <span class="mono">${e(item.client_version ? `v${item.client_version.replace(/^v/i, "")}` : "网页端")}</span>
              <small>${e(item.log_summary || "基础信息")} · ${e(formatBytes(item.log_size))}</small>
            </div>
          </td>
          <td class="mono">${e(deadline(item.created_at))}</td>
          <td>${badge(item.status)}</td>
          <td>
            <div class="feedback-row-actions">
              <button type="button" class="text-button" data-act="feedback-view" data-id="${e(item.id)}">查看详情</button>
              <a class="text-button" href="/api/v1/feedback/${encodeURIComponent(item.id)}/log" download data-stop-prop>下载日志</a>
            </div>
          </td>
        </tr>`).join("")}
      </tbody>
    </table></div>
    ${pagination("feedback", f.offset || 0, PAGE_SIZE, items.length)}`
    : emptyState(
      "暂无问题反馈记录",
      "使用过程中如果遇到报错、卡顿或排版问题，可随时点击右上角「一键上传当前日志」。",
      "bug",
      button("一键上传当前日志", "feedback-quick-upload", { cls: "primary", icon: "upload" })
    )}
  </section>`;
}

export async function openFeedbackDetail(reportId) {
  const item = await api("/feedback/" + encodeURIComponent(reportId));
  const canManage = isManager();
  modal(`反馈详情 · #${item.id.slice(0, 8)}`, `<div class="stack feedback-detail">
    <div class="feedback-detail-meta">
      <div class="feedback-reporter-card">
        <span class="avatar">${e((item.reporter_name || "?").slice(-2))}</span>
        <div class="feedback-reporter-meta">
          <strong>对应人：${e(item.reporter_name)} <span class="mono hint">(${e(item.reporter_id)})</span></strong>
          <small>${e(item.reporter_department || "未分配部门")} · ${e(roles[item.reporter_role] || item.reporter_role || "成员")}${item.reporter_email ? ` · ${e(item.reporter_email)}` : ""}</small>
        </div>
        ${badge(item.status)}
      </div>
      <div class="feedback-kv-grid">
        <div><span>问题分类</span><strong>${e(item.category_label)}</strong></div>
        <div><span>提交时间</span><strong class="mono">${e(deadline(item.created_at))}</strong></div>
        <div><span>客户端 / 服务端</span><strong class="mono">${e(item.client_version || "网页端")} / ${e(item.server_version || "—")}</strong></div>
        <div><span>运行环境</span><strong class="mono">${e(item.platform || "—")}</strong></div>
        <div><span>页面路由</span><strong class="mono">${e(item.route || "—")}</strong></div>
        <div><span>日志摘要</span><strong>${e(item.log_summary)}（${e(formatBytes(item.log_size))}）</strong></div>
      </div>
    </div>
    <section class="stack-sm">
      <h3 class="group-title">问题标题</h3>
      <p class="prose"><strong>${e(item.title)}</strong></p>
      ${item.error_message ? `<h3 class="group-title">捕获的错误信息</h3><pre class="feedback-err-box mono">${e(item.error_message)}</pre>` : ""}
      ${item.description ? `<h3 class="group-title">用户补充说明</h3><p class="prose">${e(item.description)}</p>` : ""}
      ${item.conversation_id ? `<p class="hint">关联会话：${e(item.conversation_title || item.conversation_id)} <span class="mono">(${e(item.conversation_id)})</span></p>` : ""}
    </section>
    ${canManage ? `<form class="panel pad stack-sm" data-form="feedback-update" data-id="${e(item.id)}">
      <div class="between"><h3 class="group-title">管理员处理</h3>${item.handled_by_name ? `<small class="hint">最近处理：${e(item.handled_by_name)} · ${e(deadline(item.handled_at))}</small>` : ""}</div>
      <div class="form-grid">
        <label class="field">处理状态
          <select name="status">
            ${Object.entries(feedbackStatuses).map(([k, v]) => `<option value="${e(k)}" ${item.status === k ? "selected" : ""}>${e(v)}</option>`).join("")}
          </select>
        </label>
        <label class="field span">处理备注 / 排查结论
          <textarea name="admin_note" rows="2" maxlength="4000" placeholder="记录定位结果、修复版本号或回复说明…">${e(item.admin_note || "")}</textarea>
        </label>
      </div>
      <div class="actions"><button type="submit" class="button primary sm">${icon("checkPlain")}保存处理状态</button></div>
    </form>` : (item.admin_note ? `<div class="panel pad stack-sm"><h3 class="group-title">管理员回复${item.handled_by_name ? `（${e(item.handled_by_name)}）` : ""}</h3><p class="prose">${e(item.admin_note)}</p></div>` : "")}
    <section class="stack-sm">
      <div class="between">
        <h3 class="group-title">脱敏诊断日志</h3>
        <div class="actions">
          <button type="button" class="button sm" data-act="feedback-copy-log">${icon("copy")}复制完整日志</button>
          <a class="button sm primary" href="/api/v1/feedback/${encodeURIComponent(item.id)}/log" download>${icon("download")}下载 .log 文件</a>
        </div>
      </div>
      <pre class="feedback-log-viewer mono" data-feedback-log-text>${e(item.log_text || "暂无日志")}</pre>
    </section>
  </div>`);
}

export async function saveFeedbackUpdate(form) {
  const id = form.dataset.id;
  if (!id) throw new Error("缺少反馈编号");
  const data = new FormData(form);
  const updated = await api("/feedback/" + encodeURIComponent(id), {
    method: "PATCH",
    body: {
      status: String(data.get("status") || "open"),
      admin_note: String(data.get("admin_note") || "")
    }
  });
  toastMessage(`已更新为「${updated.status_label}」`, "success");
  dialog.close();
  return updated;
}

export async function copyFeedbackLog() {
  const pre = dialog.querySelector("[data-feedback-log-text]");
  if (!pre?.textContent) return;
  await copyText(pre.textContent, "已复制完整诊断日志");
}

export function feedbackErrorStrip(errorMessage = "", opts = {}) {
  const safeErr = cleanErrorMessage(errorMessage, "");
  const source = opts.source || "chat";
  const convId = opts.conversationId || "";
  return `<div class="feedback-inline-bar">
    <button type="button" class="button sm feedback-quick-btn" data-act="feedback-quick-upload" data-source="${e(source)}" data-conversation-id="${e(convId)}" data-error="${e(safeErr.slice(0, 600))}" title="一键打包并上传桌面端与当前会话日志，自动记录对应人">${icon("upload")}<span>一键上传日志</span></button>
    <button type="button" class="text-button" data-act="feedback-open" data-source="${e(source)}" data-conversation-id="${e(convId)}" data-error="${e(safeErr.slice(0, 600))}">补充说明反馈</button>
  </div>`;
}
