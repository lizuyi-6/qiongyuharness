import { active, api, badge, button, cleanErrorMessage, copyText, deadline, dialog, e, fileMarkup, go, icon, materializeFile, modal, outputWarnings, post, readDraft, saveDraft,
  state, toastMessage, uploadMarkup, uploadOneFile, uploadSelected, uuid } from "./ui.js";
import { deckFiles, deckPanel, refreshCanvasLive, selectDeckFile, selectionChip, selectionSummaryText, setCanvasLiveState, studioArtifacts } from "./slides.js";
import { renderMarkdown } from "./markdown.js";
import { feedbackErrorStrip, recordClientLog } from "./feedback.js";
import { CATEGORY_META, openBindProjectModal, openQuickProjectContextModal } from "./project.js";

const readable = new Set(["docx", "pptx", "xlsx", "pdf", "txt", "md", "csv"]);
const imageExts = new Set(["png", "jpg", "jpeg", "webp"]);
const modelDraftKey = "chat/last-model";
const PIPELINE_MODEL = "qo:pipeline";
const imageAutoPreferenceKey = "qo/chat/allow-agent-images";
const MODEL_CACHE_TTL_MS = 5 * 60 * 1000;
let usageRequestId = 0;

// 按 threadId 跟踪正在运行的对话/生图任务，避免切换页面销毁 DOM 后丢失进度与锁定状态。
const activeTurns = new Map();
// 按 threadId 跟踪思考过程中用户追加的引导消息队列。
const queuedGuidance = new Map();
// 按 draftKey 跟踪当前正在上传 / 上传失败的附件占位项，确保拖入即显示状态且可随时取消。
const pendingUploads = new Map();
const modelCache = new Map();
const modelInFlight = new Map();
const sessionCache = new Map();
const taskCache = new Map();
const projectCache = new Map();
const titleInFlight = new Set();
const titleAttempted = new Set();
const RAW_TITLE_RE = /^(?:新会话|请?帮我|请?给我|请你|请把|做一[个份套]|写一[个份篇]|生成一[个份套]|设计一[个份套]|介绍一?下)/;

function looksLikeRawTitle(title) {
  const t = String(title || "").trim();
  if (!t || t === "新会话") return true;
  if (RAW_TITLE_RE.test(t)) return true;
  return false;
}

function applyConversationTitleLocally(threadId, nextTitle) {
  if (!threadId || !nextTitle) return;
  if (Array.isArray(state.conversations)) {
    state.conversations = state.conversations.map((item) =>
      item.id === threadId ? { ...item, title: nextTitle } : item
    );
  }
  const cached = sessionCache.get(threadId);
  if (cached) cached.title = nextTitle;
  if (state.chatWorkspace?.session?.id === threadId) {
    state.chatWorkspace.session.title = nextTitle;
    if (Array.isArray(state.conversations)) state.chatWorkspace.conversations = state.conversations;
    const titleWrap = document.querySelector("[data-thread-title-wrap]");
    if (titleWrap) {
      titleWrap.classList.remove("muted");
      titleWrap.innerHTML = `${icon("message")}<span class="thread-title-text" title="${e(nextTitle)}">${e(nextTitle)}</span>`;
    }
  }
  emitConversationsUpdated();
}

export async function generateConversationTitle(threadId, { force = false } = {}) {
  const userId = state.me?.id;
  if (!threadId || !userId || typeof window.officeDesktop?.generateTitle !== "function") return null;
  if (titleInFlight.has(threadId)) return null;
  if (!force && titleAttempted.has(threadId)) return null;
  titleInFlight.add(threadId);
  titleAttempted.add(threadId);
  try {
    const prep = await post(`/conversations/${encodeURIComponent(threadId)}/title/prepare`, {});
    if (state.me?.id !== userId) return null;
    const res = await window.officeDesktop.generateTitle({ system: prep.system, prompt: prep.prompt });
    const title = String(res?.title || "").trim();
    if (!title || title.length < 2 || state.me?.id !== userId) return null;
    const updated = await api(`/conversations/${encodeURIComponent(threadId)}/title`, {
      method: "PATCH",
      body: { title }
    });
    if (state.me?.id !== userId) return null;
    const finalTitle = updated?.title || title;
    applyConversationTitleLocally(threadId, finalTitle);
    return finalTitle;
  } catch {
    return null;
  } finally {
    titleInFlight.delete(threadId);
  }
}

export async function autoTitleRecentConversations({ force = false, limit = 12, triggerButton = null } = {}) {
  if (typeof window.officeDesktop?.generateTitle !== "function") {
    if (force) toastMessage("请在桌面客户端中使用 AI 智能命名", "info");
    return 0;
  }
  const list = Array.isArray(state.conversations) ? state.conversations : [];
  const targets = list
    .filter((s) => !s.assignment_id && !s.project_id && (force || (looksLikeRawTitle(s.title) && !titleAttempted.has(s.id))))
    .slice(0, limit);
  if (!targets.length) {
    if (force) toastMessage("近期会话标题均已清晰规范", "info");
    return 0;
  }
  if (triggerButton?.isConnected) {
    triggerButton.disabled = true;
    triggerButton.classList.add("is-spinning");
  }
  let renamed = 0;
  try {
    for (const session of targets) {
      if (!state.me?.id) break;
      const result = await generateConversationTitle(session.id, { force });
      if (result) renamed += 1;
    }
    if (force) {
      toastMessage(renamed ? `已为 ${renamed} 个会话生成 AI 标题` : "暂未能生成新标题，请稍后重试", renamed ? "success" : "info");
    }
  } finally {
    if (triggerButton?.isConnected) {
      triggerButton.disabled = false;
      triggerButton.classList.remove("is-spinning");
    }
  }
  return renamed;
}

export function runningConversationIds() {
  const userId = state.me?.id;
  if (!userId) return [];
  const ids = [];
  for (const [threadId, live] of activeTurns.entries()) {
    if (live.userId === userId) ids.push(threadId);
  }
  return ids;
}

export function resetChatState() {
  for (const live of activeTurns.values()) {
    try { void live.cancel?.(); } catch {}
  }
  for (const list of pendingUploads.values()) {
    for (const item of list) {
      try { item.controller?.abort(); } catch {}
    }
  }
  activeTurns.clear();
  queuedGuidance.clear();
  pendingUploads.clear();
  modelCache.clear();
  modelInFlight.clear();
  sessionCache.clear();
  taskCache.clear();
  projectCache.clear();
  titleInFlight.clear();
  titleAttempted.clear();
}

function activeTurnFor(threadId) {
  if (!threadId) return null;
  const live = activeTurns.get(threadId);
  if (!live || live.userId !== state.me?.id) return null;
  return live;
}

function otherActiveDesktopTurn(threadId) {
  if (!clientExecutionRequired()) return null;
  const userId = state.me?.id;
  for (const live of activeTurns.values()) {
    if (live.userId === userId && live.threadId !== threadId && live.kind === "turn") return live;
  }
  return null;
}

function isViewingThread(threadId, userId = state.me?.id) {
  if (!threadId || !userId || state.me?.id !== userId || state.route[0] !== "chat") return false;
  if (state.chatWorkspace?.session?.id === threadId) return true;
  return state.route[1] === threadId;
}

function findLiveProgress(threadId) {
  if (!threadId) return null;
  for (const node of document.querySelectorAll("[data-live-thread]")) {
    if (node.dataset?.liveThread === threadId) return node;
  }
  return null;
}

function emitConversationsUpdated() {
  try { window.dispatchEvent(new CustomEvent("qo:conversations-updated")); } catch {}
}

async function loadCheckpointFor(threadId, userId = state.me?.id) {
  if (!threadId || typeof window.officeDesktop?.getCheckpoint !== "function") return null;
  try {
    const summary = await window.officeDesktop.getCheckpoint(threadId, userId || "");
    return summary && summary.conversationId === threadId ? summary : null;
  } catch {
    return null;
  }
}

function syncCheckpointCardDom(workspace) {
  if (!workspace || state.chatWorkspace !== workspace) return;
  const turnsEl = document.querySelector("[data-chat-turns]");
  if (!turnsEl) return;
  turnsEl.querySelector("[data-checkpoint-card]")?.remove();
  const html = checkpointCardMarkup(workspace);
  if (html) {
    turnsEl.insertAdjacentHTML("beforeend", html);
    scrollToLatest();
  }
}

function sessionSignature(session) {
  if (!session) return "";
  return JSON.stringify([
    session.id,
    session.title,
    session.updated_at,
    session.has_earlier,
    (session.turns || []).map((t) => [t.id, t.status, t.reply?.length || 0, t.error || "", t.steps?.length || 0, t.file?.id || "", t.images?.length || 0, t.materials?.length || 0])
  ]);
}

// 默认允许 AI 按需自动生图（不设张数上限，按张计费），员工可随时手动关闭；按账号分开记忆。
function imageAutoPreferenceStorageKey() {
  return state.me?.id ? `${imageAutoPreferenceKey}/${state.me.id}` : "";
}

export function imageAutoPreference() {
  try {
    const key = imageAutoPreferenceStorageKey();
    return !key || localStorage.getItem(key) !== "false";
  } catch {
    return true;
  }
}

export function saveImageAutoPreference(value) {
  try {
    const key = imageAutoPreferenceStorageKey();
    if (key) localStorage.setItem(key, value ? "true" : "false");
  } catch {}
}

function clientExecutionRequired() {
  return state.engine?.client_execution === true;
}

export function desktopUsageEnabled() {
  return clientExecutionRequired() && typeof window.officeDesktop?.usage === "function";
}

function formatUsageAmount(value, unit = "USD") {
  const num = Number(value);
  if (!Number.isFinite(num)) return "—";
  const small = num !== 0 && Math.abs(num) < 0.01;
  try {
    return new Intl.NumberFormat(undefined, {
      style: "currency",
      currency: unit,
      minimumFractionDigits: 2,
      maximumFractionDigits: small ? 4 : 2
    }).format(num);
  } catch {
    return `${num.toFixed(small ? 4 : 2)} ${unit}`;
  }
}

function formatCompactAmount(value, unit = "USD") {
  const num = Number(value);
  if (!Number.isFinite(num)) return "—";
  const symbol = unit === "USD" ? "$" : unit === "CNY" || unit === "RMB" ? "¥" : "";
  const abs = Math.abs(num);
  const digits = abs !== 0 && abs < 0.01 ? 4 : Number.isInteger(num) ? 0 : 2;
  const formatted = num.toFixed(digits);
  return symbol ? `${symbol}${formatted}` : `${formatted} ${unit}`;
}

function formatUsageTime(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return "窗口未开始";
  return date.toLocaleString();
}

function formatCompactReset(value, fallback = "窗口未开始") {
  if (!value) return fallback;
  const date = new Date(value);
  const ts = date.getTime();
  if (!Number.isFinite(ts)) return fallback;
  const diffMs = ts - Date.now();
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  if (diffMs <= 60_000) return "即将重置";
  const mins = Math.ceil(diffMs / 60_000);
  if (mins < 60) return `${mins} 分钟后重置`;
  const hours = Math.floor(mins / 60);
  if (hours < 24) return `${hours} 小时后重置 · ${hh}:${mm}`;
  const month = date.getMonth() + 1;
  const day = date.getDate();
  return `${month}/${day} ${hh}:${mm} 重置`;
}

function usageRows(usage) {
  const rows = [];
  if (usage.mode === "quota_limited") {
    const labels = { "5h": "5 小时", "1d": "日", "7d": "周" };
    for (const item of usage.rate_limits || []) {
      if (item.limit <= 0) continue;
      rows.push({
        label: `${labels[item.window] || item.window}额度`, used: item.used, limit: item.limit,
        remaining: item.remaining, resetAt: item.reset_at || "", resetFallback: "窗口未开始"
      });
    }
    if (usage.quota && usage.quota.limit > 0) rows.push({
      label: "总额度", used: usage.quota.used, limit: usage.quota.limit,
      remaining: usage.quota.remaining, resetAt: "", resetFallback: "长期有效"
    });
  } else if (usage.subscription) {
    const subscription = usage.subscription;
    const plans = [
      ["日额度", "daily_usage_usd", "daily_limit_usd", ""],
      ["周额度", "weekly_usage_usd", "weekly_limit_usd", subscription.weekly_window_start ?
        new Date(Date.parse(subscription.weekly_window_start) + 7 * 24 * 60 * 60 * 1000).toISOString() : ""],
      ["月额度", "monthly_usage_usd", "monthly_limit_usd", ""]
    ];
    for (const [label, usedKey, limitKey, resetAt] of plans) {
      const limit = subscription[limitKey];
      const used = subscription[usedKey];
      if (typeof limit === "number" && limit > 0 && typeof used === "number") {
        rows.push({
          label, used, limit, remaining: Math.max(0, limit - used), resetAt,
          resetFallback: "周期重置"
        });
      }
    }
  }
  return rows;
}

function usageTone(ratio) {
  if (ratio >= 92) return "danger";
  if (ratio >= 75) return "warn";
  return "normal";
}

function renderSidebarQuota(slot, result, { loading = false } = {}) {
  if (!slot) return;
  if (!desktopUsageEnabled()) {
    slot.hidden = true;
    slot.innerHTML = "";
    return;
  }
  slot.hidden = false;
  const ringCircumference = 91.11;

  const refreshBtn = `<button type="button" class="quota-icon-btn ${loading ? "is-spinning" : ""}" data-act="usage-refresh" title="刷新模型额度" aria-label="刷新模型额度" ${loading ? "disabled" : ""}>${icon("refresh")}</button>`;

  if (loading && !result) {
    slot.innerHTML = `<section class="quota-dock is-loading" aria-label="模型额度">
      <div class="quota-dock-full">
        <div class="quota-dock-head">
          <span class="quota-dock-title"><i class="quota-dot" aria-hidden="true"></i><span>模型额度</span></span>
          <div class="quota-dock-tools">${refreshBtn}</div>
        </div>
        <div class="quota-skeleton-row"></div>
        <div class="quota-track"><span class="quota-fill is-skeleton"></span></div>
      </div>
      <button type="button" class="quota-dock-compact" data-act="usage-refresh" title="正在获取模型额度…" aria-label="刷新模型额度" disabled>
        <span class="quota-ring-label">…</span>
      </button>
    </section>`;
    return;
  }

  if (result?.state !== "available" || !result.usage) {
    const messages = {
      unauthorized: "凭证已失效，请重新登录",
      blocked: "当前网络受限，请重试或重新登录",
      unavailable: "暂无法获取额度",
      invalid: "额度数据异常"
    };
    const text = messages[result?.state] || messages.unavailable;
    slot.innerHTML = `<section class="quota-dock is-muted" aria-label="模型额度">
      <div class="quota-dock-full">
        <div class="quota-dock-head">
          <span class="quota-dock-title"><i class="quota-dot" aria-hidden="true"></i><span>模型额度</span></span>
          <div class="quota-dock-tools">${refreshBtn}</div>
        </div>
        <p class="quota-dock-meta" role="status">${e(text)}</p>
      </div>
      <button type="button" class="quota-dock-compact" data-act="usage-refresh" title="${e(text)}（点击重试）" aria-label="刷新模型额度">
        <span class="quota-ring-label">—</span>
      </button>
    </section>`;
    return;
  }

  const usage = result.usage;
  const unit = usage.unit || "USD";

  if (usage.balance !== undefined || usage.planName === "钱包余额") {
    const val = usage.balance ?? usage.remaining;
    const compact = formatCompactAmount(val, unit);
    const full = formatUsageAmount(val, unit);
    slot.innerHTML = `<section class="quota-dock is-normal" aria-label="钱包余额">
      <div class="quota-dock-full">
        <div class="quota-dock-head">
          <span class="quota-dock-title"><i class="quota-dot" aria-hidden="true"></i><span>钱包余额</span></span>
          <div class="quota-dock-tools">${refreshBtn}</div>
        </div>
        <div class="quota-dock-metrics">
          <strong class="quota-dock-rem" title="可用余额 ${e(full)}">${e(compact)}</strong>
          <span class="quota-dock-pill">可用</span>
        </div>
      </div>
      <button type="button" class="quota-dock-compact" data-act="usage-refresh" title="钱包余额：${e(full)}（点击刷新）" aria-label="刷新模型额度">
        <span class="quota-ring-label">${e(compact)}</span>
      </button>
    </section>`;
    return;
  }

  const rows = usageRows(usage);
  if (!rows.length) {
    slot.innerHTML = `<section class="quota-dock is-normal" aria-label="模型额度">
      <div class="quota-dock-full">
        <div class="quota-dock-head">
          <span class="quota-dock-title"><i class="quota-dot" aria-hidden="true"></i><span>模型额度</span></span>
          <div class="quota-dock-tools">${refreshBtn}</div>
        </div>
        <div class="quota-dock-metrics">
          <strong class="quota-dock-rem">不限额</strong>
          <span class="quota-dock-pill">正常</span>
        </div>
      </div>
      <button type="button" class="quota-dock-compact" data-act="usage-refresh" title="当前账户不限额（点击刷新）" aria-label="刷新模型额度">
        <span class="quota-ring-label">∞</span>
      </button>
    </section>`;
    return;
  }

  const primary = rows[0];
  const extras = rows.slice(1);
  const ratio = Math.max(0, Math.min(100, (primary.used / primary.limit) * 100));
  const remainPct = Math.max(0, Math.min(100, Math.round(100 - ratio)));
  const tone = usageTone(ratio);
  const dashOffset = (ringCircumference * (1 - remainPct / 100)).toFixed(2);
  const expanded = Boolean(state.usageOpen && extras.length);
  const remCompact = primary.remaining <= 0 ? "已用尽" : formatCompactAmount(primary.remaining, unit);
  const usedSummary = `${formatCompactAmount(primary.used, unit)} / ${formatCompactAmount(primary.limit, unit)}`;
  const resetCompact = formatCompactReset(primary.resetAt, primary.resetFallback);
  const resetFull = primary.resetAt ? `重置时间：${formatUsageTime(primary.resetAt)}` : primary.resetFallback;
  const fullTooltip = `${primary.label}：剩余 ${formatUsageAmount(primary.remaining, unit)}（已用 ${formatUsageAmount(primary.used, unit)} / ${formatUsageAmount(primary.limit, unit)}）\n${resetFull}`;

  const extrasHtml = extras.length ? `<div class="quota-dock-extra" ${expanded ? "" : "hidden"}>
    ${extras.map((row) => {
      const r = Math.max(0, Math.min(100, (row.used / row.limit) * 100));
      const t = usageTone(r);
      const rTip = `${row.label}：剩余 ${formatUsageAmount(row.remaining, unit)}（已用 ${formatUsageAmount(row.used, unit)} / ${formatUsageAmount(row.limit, unit)}）`;
      return `<div class="quota-mini-row is-${t}" title="${e(rTip)}">
        <div class="quota-mini-top">
          <span>${e(row.label)}</span>
          <strong class="mono">${e(row.remaining <= 0 ? "已用尽" : formatCompactAmount(row.remaining, unit))}</strong>
        </div>
        <progress class="quota-progress" max="100" value="${Math.round(r)}" aria-label="${e(row.label)}使用比例"></progress>
      </div>`;
    }).join("")}
  </div>` : "";

  slot.innerHTML = `<section class="quota-dock is-${tone}" aria-label="模型额度">
    <div class="quota-dock-full">
      <div class="quota-dock-head">
        <span class="quota-dock-title" title="${e(fullTooltip)}">
          <i class="quota-dot" aria-hidden="true"></i>
          <span>${e(primary.label)}</span>
        </span>
        <div class="quota-dock-tools">
          ${extras.length ? `<button type="button" class="quota-icon-btn quota-expand-btn ${expanded ? "is-open" : ""}" data-act="usage-toggle" aria-expanded="${expanded ? "true" : "false"}" title="${expanded ? "收起其他周期" : `展开其余 ${extras.length} 项额度`}">
            <span class="quota-more-count">+${extras.length}</span>${icon("down")}
          </button>` : ""}
          ${refreshBtn}
        </div>
      </div>
      <div class="quota-dock-metrics" title="${e(fullTooltip)}">
        <strong class="quota-dock-rem">${e(remCompact)}</strong>
        <span class="quota-dock-used">${e(usedSummary)}</span>
      </div>
      <progress class="quota-progress" max="100" value="${Math.round(ratio)}" aria-label="${e(primary.label)}使用比例"></progress>
      <p class="quota-dock-meta" title="${e(resetFull)}">${e(resetCompact)}</p>
      ${extrasHtml}
    </div>
    <button type="button" class="quota-dock-compact" data-act="usage-refresh" title="${e(fullTooltip)}&#10;点击刷新" aria-label="刷新模型额度">
      <svg class="quota-ring" viewBox="0 0 36 36" aria-hidden="true">
        <circle class="quota-ring-track" cx="18" cy="18" r="14.5"></circle>
        <circle class="quota-ring-fill" cx="18" cy="18" r="14.5" stroke-dasharray="${ringCircumference}" stroke-dashoffset="${dashOffset}"></circle>
      </svg>
      <span class="quota-ring-label">${remainPct}%</span>
    </button>
  </section>`;
}

export async function refreshDesktopUsage(workspace = state.chatWorkspace, slot = document.querySelector("[data-sidebar-quota]")) {
  if (!desktopUsageEnabled()) return null;
  const requestId = ++usageRequestId;
  renderSidebarQuota(slot, state.usageState, { loading: true });
  let result;
  try { result = await window.officeDesktop.usage(); }
  catch { result = { state: "unavailable" }; }
  if (requestId !== usageRequestId) return result;
  state.usageState = result;
  if (workspace) workspace.usageState = result;
  const currentSlot = document.querySelector("[data-sidebar-quota]");
  if (currentSlot) renderSidebarQuota(currentSlot, result, { loading: false });
  return result;
}

export function syncDesktopUsagePanel() {
  const slot = document.querySelector("[data-sidebar-quota]");
  if (!slot) return;
  if (!desktopUsageEnabled()) {
    slot.hidden = true;
    slot.innerHTML = "";
    return;
  }
  if (state.usageState) {
    renderSidebarQuota(slot, state.usageState, { loading: false });
  } else {
    void refreshDesktopUsage(state.chatWorkspace, slot);
  }
}

export async function toggleDesktopUsage() {
  state.usageOpen = !state.usageOpen;
  syncDesktopUsagePanel();
}

export function syncShellCanvasState() {
  const shellEl = document.querySelector(".app-shell");
  if (!shellEl) return;
  const isChat = state.route[0] === "chat";
  const sideBySide = isChat && Boolean(state.contextOpen) && state.studioTab === "deck";
  shellEl.classList.toggle("canvas-side-by-side", sideBySide);
}

function syncStudioDom() {
  const workspaceEl = document.querySelector(".agent-workspace");
  if (!workspaceEl) return;
  const open = Boolean(state.contextOpen);
  const tab = state.studioTab === "deck" ? "deck" : "context";
  state.studioTab = tab;
  workspaceEl.classList.toggle("studio-open", open);
  workspaceEl.classList.toggle("studio-deck", open && tab === "deck");

  for (const pane of workspaceEl.querySelectorAll("[data-studio-pane]")) {
    pane.hidden = pane.dataset.studioPane !== tab;
  }
  for (const btn of workspaceEl.querySelectorAll("[data-act=studio-tab]")) {
    const activeTab = btn.dataset.id === tab;
    btn.classList.toggle("active", activeTab);
    btn.setAttribute("aria-selected", String(activeTab));
  }
  for (const toggleBtn of workspaceEl.querySelectorAll("[data-act=toggle-context]")) {
    const targetTab = toggleBtn.dataset.id || "context";
    const isOn = open && tab === targetTab;
    toggleBtn.classList.toggle("active", isOn);
    toggleBtn.setAttribute("aria-expanded", String(isOn));
  }
  syncShellCanvasState();
}

export function toggleWorkspaceContext(triggerButton) {
  const workspaceEl = document.querySelector(".agent-workspace");
  if (!workspaceEl) return;
  const requestedTab = triggerButton?.dataset?.id || "";
  if (requestedTab === "close") {
    state.contextOpen = false;
  } else if (requestedTab && (requestedTab === "deck" || requestedTab === "context")) {
    if (state.contextOpen && state.studioTab === requestedTab) {
      state.contextOpen = false;
    } else {
      state.contextOpen = true;
      state.studioTab = requestedTab;
    }
  } else {
    state.contextOpen = !state.contextOpen;
  }
  syncStudioDom();
}

export function switchStudioTab(tab) {
  state.contextOpen = true;
  state.studioTab = tab === "deck" ? "deck" : "context";
  syncStudioDom();
}

export function openDeckInStudio(fileId) {
  state.contextOpen = true;
  state.studioTab = "deck";
  syncStudioDom();
  if (fileId) selectDeckFile(fileId);
}

function desktopModelCacheKey(allowPipeline, draftKey) {
  return `${state.me?.id || ""}|${allowPipeline ? 1 : 0}|${draftKey}`;
}

function buildDesktopModelState(raw, { allowPipeline = true, draftKey = modelDraftKey } = {}) {
  const storedModel = readDraft(draftKey)?.model || "";
  const models = [...new Set((Array.isArray(raw?.models) ? raw.models : [])
    .filter((model) => typeof model === "string" && model.trim()).map((model) => model.trim()))];
  const image_models = [...new Set((Array.isArray(raw?.image_models) ? raw.image_models : [])
    .filter((model) => typeof model === "string" && /^gpt-image-/.test(model.trim())).map((model) => model.trim()))].sort();
  const pipelineAvailable = allowPipeline && raw?.pipeline?.available === true;
  const pipelineReason = allowPipeline && typeof raw?.pipeline?.reason === "string" ? raw.pipeline.reason : "";
  const storedModelAvailable = storedModel === PIPELINE_MODEL ? pipelineAvailable : models.includes(storedModel);
  const invalidated = Boolean(storedModel && !storedModelAvailable);
  const selectedModel = invalidated ? "" : storedModel || (pipelineAvailable ? PIPELINE_MODEL : "");
  if (invalidated) saveDraft(draftKey, { model: "" });
  const modelNotice = invalidated ? "上次选择的模型已不可用，已清空，请重新选择。" :
    raw?.error || (models.length || pipelineAvailable ? "" : "暂无可用模型，请联系管理员开通。");
  return { models, image_models, selectedModel, modelNotice, pipelineAvailable, pipelineReason, modelsLoading: false };
}

function cachedDesktopModelState({ allowPipeline = true, draftKey = modelDraftKey } = {}) {
  if (!clientExecutionRequired()) return { state: {}, fresh: true };
  if (typeof window.officeDesktop?.runAgent !== "function") {
    return { state: { models: [], image_models: [], selectedModel: "", modelNotice: "", modelsLoading: false }, fresh: true };
  }
  const entry = modelCache.get(desktopModelCacheKey(allowPipeline, draftKey));
  if (!entry) return null;
  const fresh = (Date.now() - entry.fetchedAt) < MODEL_CACHE_TTL_MS;
  return { state: buildDesktopModelState(entry.raw, { allowPipeline, draftKey }), fresh };
}

export async function loadDesktopModelState({ allowPipeline = true, draftKey = modelDraftKey, force = false } = {}) {
  if (!clientExecutionRequired()) return {};
  if (typeof window.officeDesktop?.runAgent !== "function") return { models: [], image_models: [], selectedModel: "", modelNotice: "" };

  const cacheKey = desktopModelCacheKey(allowPipeline, draftKey);
  if (!force) {
    const cached = cachedDesktopModelState({ allowPipeline, draftKey });
    if (cached?.fresh) return cached.state;
  }
  if (!force && modelInFlight.has(cacheKey)) {
    return modelInFlight.get(cacheKey);
  }
  const job = (async () => {
    try {
      if (typeof window.officeDesktop.models !== "function") throw new Error("当前桌面客户端不支持模型列表");
      const result = await window.officeDesktop.models();
      const built = buildDesktopModelState(result, { allowPipeline, draftKey });
      if (built.models.length || built.pipelineAvailable) {
        modelCache.set(cacheKey, { raw: result, fetchedAt: Date.now() });
      }
      return built;
    } catch (error) {
      return { models: [], image_models: [], selectedModel: "", modelNotice: error.message || "模型列表获取失败，请重试。", modelsLoading: false };
    } finally {
      modelInFlight.delete(cacheKey);
    }
  })();
  modelInFlight.set(cacheKey, job);
  return job;
}

export function prefetchDesktopModels() {
  if (!clientExecutionRequired() || typeof window.officeDesktop?.models !== "function") return;
  void loadDesktopModelState().catch(() => {});
}

const MODE_OPTIONS = [
  { id: "auto", label: "自动", hint: "按需回复或导出文件", glyph: "sparkles" },
  { id: "reply", label: "仅回复", hint: "对话内直接 Markdown 回复", glyph: "message" },
  { id: "docx", label: "Word", hint: "导出可编辑 Word 文档", glyph: "docx" },
  { id: "pptx", label: "PPT", hint: "AI 自由画布设计演示文稿", glyph: "pptx" }
];

function modeOptionFor(value) {
  return MODE_OPTIONS.find((item) => item.id === value) || MODE_OPTIONS[0];
}

function modePickerInner(currentMode, disabled = false) {
  const activeMode = modeOptionFor(currentMode);
  return `<input type="hidden" name="mode" value="${e(activeMode.id)}">
    <button type="button" class="composer-picker-trigger" data-act="picker-toggle" data-id="mode" aria-haspopup="listbox" aria-expanded="false" title="输出方式：${e(activeMode.label)}" ${disabled ? "disabled" : ""}>
      <span class="picker-glyph">${icon(activeMode.glyph)}</span>
      <span class="picker-label">${e(activeMode.label)}</span>
      <span class="picker-caret">${icon("down")}</span>
    </button>
    <div class="composer-picker-menu" data-picker-menu="mode" role="listbox" aria-label="输出方式" hidden>
      ${MODE_OPTIONS.map((item) => {
        const on = item.id === activeMode.id;
        return `<button type="button" class="composer-picker-item ${on ? "active" : ""}" role="option" aria-selected="${on ? "true" : "false"}" data-act="picker-pick" data-picker="mode" data-id="${e(item.id)}">
          <span class="picker-item-glyph">${icon(item.glyph)}</span>
          <span class="picker-item-copy"><strong>${e(item.label)}</strong><small>${e(item.hint)}</small></span>
          <span class="picker-item-check">${on ? icon("checkPlain") : ""}</span>
        </button>`;
      }).join("")}
    </div>`;
}

function modelPickerInner(workspace, disabled = false) {
  const models = workspace?.models || [];
  const selected = workspace?.selectedModel || "";
  const isPipeline = selected === PIPELINE_MODEL;
  const hasAny = Boolean(models.length || workspace?.pipelineAvailable);
  const label = isPipeline
    ? "智能分工"
    : selected
      ? selected
      : workspace?.modelsLoading
        ? "加载模型…"
        : models.length
          ? "选择模型"
          : "暂无模型";
  const glyph = isPipeline ? "sparkles" : "sliders";
  const triggerDisabled = disabled || !hasAny;
  return `<button type="button" class="composer-picker-trigger model-picker-trigger ${isPipeline ? "is-pipeline" : ""}" data-chat-model data-act="picker-toggle" data-id="model" aria-haspopup="listbox" aria-expanded="false" title="执行模型：${e(label)}" ${triggerDisabled ? "disabled" : ""}>
      <span class="picker-glyph">${icon(glyph)}</span>
      <span class="picker-label ${!isPipeline && selected ? "mono" : ""}">${e(label)}</span>
      <span class="picker-caret">${icon("down")}</span>
    </button>
    <div class="composer-picker-menu align-right" data-picker-menu="model" role="listbox" aria-label="选择执行模型" hidden>
      <button type="button" class="composer-picker-item ${isPipeline ? "active" : ""}" role="option" aria-selected="${isPipeline ? "true" : "false"}" data-act="picker-pick" data-picker="model" data-id="${PIPELINE_MODEL}" ${workspace?.pipelineAvailable ? "" : "disabled"}>
        <span class="picker-item-glyph">${icon("sparkles")}</span>
        <span class="picker-item-copy"><strong>智能分工</strong><small>${e(workspace?.pipelineAvailable ? "理解 · 执行 · 验收协同" : (workspace?.pipelineReason || "暂不可用"))}</small></span>
        <span class="picker-item-check">${isPipeline ? icon("checkPlain") : ""}</span>
      </button>
      ${models.length ? '<div class="composer-picker-divider" aria-hidden="true"></div>' : ""}
      ${models.map((model) => {
        const on = selected === model;
        return `<button type="button" class="composer-picker-item ${on ? "active" : ""}" role="option" aria-selected="${on ? "true" : "false"}" data-act="picker-pick" data-picker="model" data-id="${e(model)}">
          <span class="picker-item-glyph">${icon("sliders")}</span>
          <span class="picker-item-copy"><strong class="mono">${e(model)}</strong></span>
          <span class="picker-item-check">${on ? icon("checkPlain") : ""}</span>
        </button>`;
      }).join("")}
    </div>`;
}

export function closeComposerPickers() {
  for (const picker of document.querySelectorAll(".composer-picker.is-open")) {
    picker.classList.remove("is-open");
    picker.querySelector("[data-picker-menu]")?.setAttribute("hidden", "");
    picker.querySelector("[data-act=picker-toggle]")?.setAttribute("aria-expanded", "false");
  }
}

export function toggleComposerPicker(name) {
  const form = chatForm();
  if (!form || form.dataset.busy) return;
  const picker = form.querySelector(`.composer-picker[data-picker="${name}"]`);
  if (!picker) return;
  const trigger = picker.querySelector("[data-act=picker-toggle]");
  const menu = picker.querySelector("[data-picker-menu]");
  if (!trigger || trigger.disabled || !menu) return;
  const opening = menu.hidden;
  closeSlashMenu();
  closeComposerPickers();
  if (opening) {
    picker.classList.add("is-open");
    menu.hidden = false;
    trigger.setAttribute("aria-expanded", "true");
  }
}

function syncModePickerDom(form) {
  const picker = form?.querySelector('.composer-picker[data-picker="mode"]');
  if (!picker) return;
  const current = form.elements?.mode?.value || "auto";
  picker.innerHTML = modePickerInner(current, Boolean(form.dataset.busy));
}

export function pickComposerOption(pickerName, value) {
  const form = chatForm();
  const workspace = state.chatWorkspace;
  closeComposerPickers();
  if (!form || !workspace || form.dataset.busy) return;
  if (pickerName === "mode") {
    const next = modeOptionFor(value).id;
    if (form.elements.mode) form.elements.mode.value = next;
    syncModePickerDom(form);
    saveWorkspaceInput(form);
    return;
  }
  if (pickerName === "model") {
    const allowed = value === PIPELINE_MODEL ? workspace.pipelineAvailable === true : workspace.models?.includes(value);
    if (!allowed) return;
    workspace.selectedModel = value;
    workspace.modelNotice = "";
    saveDraft(modelDraftKey, { model: value });
    syncComposerModelsDom(workspace);
    const status = document.querySelector("[data-chat-model-status]");
    if (status) {
      status.textContent = workspace.pipelineReason || "";
      status.hidden = !workspace.pipelineReason;
    }
  }
}

function syncComposerModelsDom(workspace) {
  const form = chatForm();
  if (!form || state.chatWorkspace !== workspace) return;
  const modelPicker = form.querySelector('.composer-picker[data-picker="model"]');
  if (modelPicker) {
    modelPicker.innerHTML = modelPickerInner(workspace, Boolean(form.dataset.busy));
  }
  const statusText = [workspace.modelNotice, workspace.pipelineReason].filter(Boolean).join(" ");
  let statusEl = form.querySelector("[data-chat-model-status]");
  if (statusText) {
    if (!statusEl) {
      statusEl = document.createElement("p");
      statusEl.className = "chat-model-status";
      statusEl.dataset.chatModelStatus = "";
      statusEl.setAttribute("role", "status");
      form.append(statusEl);
    }
    statusEl.textContent = statusText;
    statusEl.hidden = false;
    statusEl.classList.toggle("error", !(workspace.models?.length || workspace.pipelineAvailable));
  } else if (statusEl) {
    statusEl.textContent = "";
    statusEl.hidden = true;
  }
  const showImageControls = clientExecutionRequired() && typeof window.officeDesktop?.runAgent === "function" && (workspace.image_models?.length > 0);
  if (showImageControls && !form.querySelector("[data-image-toggle]")) {
    const imageAutoOn = imageAutoPreference();
    const drawerOpen = Boolean(state.imageDrawerOpen);
    const sep = form.querySelector(".composer-tools .composer-sep");
    const btnHtml = `<button type="button" class="composer-tool ${drawerOpen ? "active" : ""} ${imageAutoOn ? "dot-active" : ""}" data-act="toggle-image-drawer" data-image-toggle aria-expanded="${drawerOpen ? "true" : "false"}" title="图片生成设置" aria-label="图片生成设置">${icon("image")}</button>`;
    if (sep) sep.insertAdjacentHTML("beforebegin", btnHtml);
    else form.querySelector(".composer-tools")?.insertAdjacentHTML("beforeend", btnHtml);
    const drawerMarkup = `<div class="image-drawer ${drawerOpen ? "open" : ""}" data-image-drawer aria-label="图片生成设置">
      <div class="image-drawer-inner">
        <div class="composer-image">
          <div class="image-controls">
            <label><span>尺寸</span><select name="image_size" aria-label="生图尺寸">
              <option value="1024x1024">方形 1024×1024</option><option value="1536x1024">横版 1536×1024</option>
              <option value="1024x1536">竖版 1024×1536</option><option value="1536x864">16:9 1536×864</option>
            </select></label>
            <label><span>质量</span><select name="image_quality" aria-label="生图质量">
              <option value="low">草稿</option><option value="medium" selected>标准</option><option value="high">精细</option>
            </select></label>
            ${button("生成图片", "chat-generate-image", { cls: "sm", icon: "image", disabled: Boolean(form.dataset.busy) })}
          </div>
          <label class="image-auto-option" data-image-auto-wrap><input type="checkbox" name="allow_ai_images" ${imageAutoPreference() ? "checked" : ""}>
            <span>允许 AI 按需自动生图（不限张数，按张计费）</span></label>
        </div>
      </div>
    </div>`;
    if (statusEl) statusEl.insertAdjacentHTML("beforebegin", drawerMarkup);
    else form.insertAdjacentHTML("beforeend", drawerMarkup);
  }
  syncImageAutoIndicator();
  syncChatSubmit(form, workspace);
}

function selectedModelAvailable(workspace) {
  return workspace?.selectedModel === PIPELINE_MODEL ? workspace.pipelineAvailable === true :
    Boolean(workspace?.selectedModel && workspace.models?.includes(workspace.selectedModel));
}

export function canSubmitChat(workspace = state.chatWorkspace) {
  if (!workspace) return false;
  if (workspace.task && !active(workspace.task.status)) return false;
  const desktopRequired = clientExecutionRequired();
  if (!desktopRequired) return true;
  if (typeof window.officeDesktop?.runAgent !== "function") return false;
  return selectedModelAvailable(workspace);
}

function pendingUploadListFor(draftKey) {
  if (!draftKey) return [];
  return pendingUploads.get(draftKey) || [];
}

function hasActiveUploadsFor(workspace = state.chatWorkspace) {
  const key = workspace?.draftKey;
  if (!key) return false;
  return pendingUploadListFor(key).some((item) => item.status === "uploading");
}

function readableSourcesOnly(list) {
  return (Array.isArray(list) ? list : []).filter((file) => readable.has(String(file?.extension || "").toLowerCase()));
}

function formatFileSizeCompact(bytes) {
  const n = Number(bytes);
  if (!Number.isFinite(n) || n <= 0) return "";
  if (n >= 1024 * 1024) return `${(n / (1024 * 1024)).toFixed(1)} MB`;
  return `${Math.max(1, Math.ceil(n / 1024))} KB`;
}

function sourceIconName(ext) {
  const low = String(ext || "").toLowerCase();
  if (imageExts.has(low)) return "image";
  if (low === "docx" || low === "pptx" || low === "xlsx") return low;
  return "clip";
}

export function syncChatSubmit(form = chatForm(), workspace = state.chatWorkspace) {
  if (!form) return;
  const isBusy = Boolean(form.dataset.busy);
  const uploading = hasActiveUploadsFor(workspace);
  const threadId = currentLiveThreadId(workspace);
  const live = activeTurnFor(threadId);
  const hasText = Boolean(form.elements?.message?.value?.trim());
  const canQueue = Boolean(isBusy && live && live.kind !== "image" && hasText);
  const disabled = uploading ? true : (isBusy ? !canQueue : !canSubmitChat(workspace));
  form.querySelectorAll('button[type="submit"]').forEach((control) => {
    control.disabled = disabled;
    control.classList.toggle("is-queue-mode", Boolean(isBusy && live && live.kind !== "image"));
    control.classList.toggle("ready", hasText && !disabled);
    control.title = uploading
      ? "等待文件上传完成后发送…"
      : (isBusy && live && live.kind !== "image" ? "追加引导（Enter）" : "发送（Enter）");
    control.setAttribute("aria-label", uploading ? "等待文件上传完成" : (isBusy && live && live.kind !== "image" ? "追加引导信息" : "发送消息"));
  });
  const resumeBtn = document.querySelector(".checkpoint-resume-btn");
  if (resumeBtn) {
    const canResume = !form.dataset.busy && !uploading && (canSubmitChat(workspace) || Boolean(workspace?.checkpoint?.hasCompletedResult));
    resumeBtn.disabled = !canResume;
  }
}

function currentLiveThreadId(workspace = state.chatWorkspace) {
  const direct = workspace?.session?.id
    || (state.route[0] === "chat" && state.route[1] && !["task", "project"].includes(state.route[1]) ? state.route[1] : "");
  if (direct && activeTurns.has(direct)) return direct;
  const userId = state.me?.id;
  if (direct) return direct;
  for (const [id, live] of activeTurns.entries()) {
    if (!userId || live.userId === userId) return id;
  }
  return "";
}

function guidanceListFor(threadId) {
  if (!threadId) return [];
  return queuedGuidance.get(threadId) || [];
}

function guidanceQueueMarkup(threadId) {
  const items = guidanceListFor(threadId);
  if (!items.length) return "";
  const live = activeTurnFor(threadId);
  const canRedirect = Boolean(live && live.kind !== "image" && live.canCancel && !live.cancelling);
  return `<div class="guidance-queue-card" role="region" aria-label="已追加引导队列">
    <div class="guidance-queue-head">
      <span class="guidance-queue-title">${icon("sparkles")}<strong>已追加 ${items.length} 条引导</strong><small>本轮完成后将自动发给 AI 继续处理</small></span>
      <div class="guidance-queue-actions">
        ${canRedirect ? `<button type="button" class="button sm primary" data-act="chat-guidance-apply-now" title="中止当前思考并合并原要求与追加引导立刻重做">${icon("refresh")}立即按补充重做</button>` : ""}
        <button type="button" class="text-button" data-act="chat-guidance-clear">清空</button>
      </div>
    </div>
    <ul class="guidance-queue-list">${items.map((item, idx) => `
      <li class="guidance-queue-item">
        <span class="guidance-queue-idx">${idx + 1}</span>
        <span class="guidance-queue-text">${e(item.text)}</span>
        <button type="button" class="guidance-queue-del" data-act="chat-guidance-remove" data-id="${e(item.id)}" aria-label="移除该条引导" title="移除">${icon("x")}</button>
      </li>`).join("")}
    </ul>
  </div>`;
}

function syncGuidanceDom(threadId = currentLiveThreadId()) {
  const slot = document.querySelector("[data-chat-guidance]");
  if (!slot) return;
  slot.innerHTML = guidanceQueueMarkup(threadId);
}

export function enqueueGuidanceFromForm(form = chatForm()) {
  const workspace = state.chatWorkspace;
  const textarea = form?.elements?.message;
  const text = textarea?.value?.trim() || "";
  if (!text) return false;
  const threadId = currentLiveThreadId(workspace);
  const live = activeTurnFor(threadId);
  if (!threadId || !live || live.kind === "image") {
    toastMessage("当前正在生成图片，请稍候完成后再发送", "info");
    return false;
  }
  const list = [...guidanceListFor(threadId)];
  if (list.length >= 6) {
    toastMessage("单轮最多追加 6 条引导，可先点「立即按补充重做」或精简后提交", "error");
    return false;
  }
  list.push({ id: uuid(), text, createdAt: Date.now() });
  queuedGuidance.set(threadId, list);
  textarea.value = "";
  syncComposerGrow(textarea);
  syncGuidanceDom(threadId);
  syncChatSubmit(form, workspace);
  toastMessage("已追加引导，本轮完成后自动发给 AI（也可点「立即按补充重做」）", "success");
  return true;
}

export function removeQueuedGuidance(id) {
  const threadId = currentLiveThreadId();
  if (!threadId || !id) return;
  const next = guidanceListFor(threadId).filter((item) => item.id !== id);
  if (next.length) queuedGuidance.set(threadId, next);
  else queuedGuidance.delete(threadId);
  syncGuidanceDom(threadId);
}

export function clearQueuedGuidance() {
  const threadId = currentLiveThreadId();
  if (!threadId) return;
  queuedGuidance.delete(threadId);
  syncGuidanceDom(threadId);
}

export async function applyQueuedGuidanceNow() {
  const threadId = currentLiveThreadId();
  const live = activeTurnFor(threadId);
  const items = guidanceListFor(threadId);
  if (!live || !items.length) return;
  if (!live.canCancel || typeof live.cancel !== "function") {
    toastMessage("当前执行模式不支持中途改向，将在本轮完成后自动按引导继续", "info");
    return;
  }
  live.redirectGuidance = items.map((item) => item.text);
  queuedGuidance.delete(threadId);
  syncGuidanceDom(threadId);
  toastMessage("正在按追加引导重新规划执行…", "info");
  await cancelActiveChat();
}

function formatMergedGuidanceMessage(baseMessage, guidanceTexts) {
  const lines = guidanceTexts.map((t, i) => `${i + 1}. ${t}`).join("\n");
  if (!baseMessage) return guidanceTexts.length === 1 ? guidanceTexts[0] : lines;
  return `${baseMessage}\n\n【思考过程中追加的补充引导（请严格结合以下最新要求执行）】\n${lines}`;
}

function formatFollowUpGuidanceMessage(guidanceTexts) {
  if (guidanceTexts.length === 1) return guidanceTexts[0];
  return `请结合以下补充引导继续调整完善：\n${guidanceTexts.map((t, i) => `${i + 1}. ${t}`).join("\n")}`;
}

function draftFor(workspace) {
  return readDraft(workspace.draftKey) || {};
}

function writeDraft(workspace, form, fields = {}) {
  const previous = draftFor(workspace);
  const inputs = form ? { message: form.elements.message.value, mode: form.elements.mode.value } : {};
  saveDraft(workspace.draftKey, { ...previous, ...inputs, ...fields });
}

function updateContextDom(workspace) {
  if (state.chatWorkspace !== workspace) return;
  const thread = document.querySelector(".agent-thread");
  if (thread) {
    const existingAnchor = thread.querySelector(".context-anchor");
    const anchorHtml = contextAnchorBar(workspace);
    if (anchorHtml) {
      if (existingAnchor) existingAnchor.outerHTML = anchorHtml;
      else thread.querySelector(".thread-header")?.insertAdjacentHTML("afterend", anchorHtml);
    } else {
      existingAnchor?.remove();
    }
  }
  const contextPane = document.querySelector('[data-studio-pane="context"]');
  if (contextPane) contextPane.innerHTML = contextBody(workspace);
}

export async function loadWorkspace(parts) {
  const [, scope, assignment] = parts;
  const userId = state.me?.id;
  const isSessionRoute = Boolean(scope && !["task", "project"].includes(scope));

  // 1. 桌面模型列表：命中缓存直接瞬开，未命中/过期时非阻塞后台拉取，不再卡住页面切换。
  const cachedModels = cachedDesktopModelState();
  const desktopModels = cachedModels
    ? cachedModels.state
    : {
        models: [],
        image_models: [],
        selectedModel: readDraft(modelDraftKey)?.model || "",
        modelNotice: "",
        pipelineAvailable: false,
        pipelineReason: "",
        modelsLoading: clientExecutionRequired() && typeof window.officeDesktop?.runAgent === "function"
      };
  const needModelRefresh = !cachedModels || !cachedModels.fresh;

  // 2. 会话详情 / 关联任务 / 关联项目：优先使用内存缓存瞬开，再后台静默校准。
  let session = isSessionRoute ? (sessionCache.get(scope) || null) : null;
  let sessionFromCache = Boolean(session);
  if (isSessionRoute && !session) {
    session = await api("/conversations/" + encodeURIComponent(scope));
    sessionCache.set(scope, session);
  }

  const bound = session?.assignment_id || (scope === "task" ? assignment : null);
  const group = session?.project_id || (scope === "project" ? assignment : null);
  let task = bound ? (taskCache.get(bound) || null) : null;
  let project = group ? (projectCache.get(group) || null) : null;
  const taskFromCache = Boolean(task);
  const projectFromCache = Boolean(project);

  // 仅当直接进入新任务/项目会话且本地无缓存时才等待对应实体；已有会话切回时不阻塞。
  const initialFetches = [];
  if (scope === "task" && bound && !task) {
    initialFetches.push(api("/assignments/" + encodeURIComponent(bound)).then((res) => { task = res; taskCache.set(bound, res); }));
  }
  if (scope === "project" && group && !project) {
    initialFetches.push(api("/projects/" + encodeURIComponent(group)).then((res) => { project = res; projectCache.set(group, res); }));
  }
  if (initialFetches.length) await Promise.all(initialFetches);

  const conversations = Array.isArray(state.conversations) ? state.conversations : [];
  const todayItems = Array.isArray(state.today) ? state.today : [];
  const checkpoint = session?.id ? await loadCheckpointFor(session.id, userId) : null;

  const workspace = {
    session,
    checkpoint,
    conversations,
    task,
    project,
    ...desktopModels,
    tasks: todayItems.filter((item) => active(item.status)),
    draftKey: "chat/" + (session?.id || (task ? "task/" + task.id : project ? "project/" + project.id : "new"))
  };

  workspace._revalidate = () => {
    if (workspace._revalidated) return;
    workspace._revalidated = true;

    if (needModelRefresh && clientExecutionRequired() && typeof window.officeDesktop?.runAgent === "function") {
      void loadDesktopModelState({ force: Boolean(cachedModels && !cachedModels.fresh) }).then((freshModels) => {
        if (state.me?.id !== userId || state.chatWorkspace !== workspace) return;
        Object.assign(workspace, freshModels);
        syncComposerModelsDom(workspace);
      }).catch(() => {});
    }

    // 后台刷新当前会话（若前面用的是缓存，校验是否有最新回复或状态变更）
    if (isSessionRoute && sessionFromCache) {
      void api("/conversations/" + encodeURIComponent(scope)).then(async (freshSession) => {
        if (state.me?.id !== userId) return;
        sessionCache.set(scope, freshSession);
        if (state.chatWorkspace !== workspace) return;
        const freshCheckpoint = await loadCheckpointFor(scope, userId);
        const cpChanged = JSON.stringify(workspace.checkpoint || null) !== JSON.stringify(freshCheckpoint || null);
        workspace.checkpoint = freshCheckpoint;
        const changed = sessionSignature(workspace.session) !== sessionSignature(freshSession);
        workspace.session = freshSession;
        if (changed || cpChanged) {
          const turnsEl = document.querySelector("[data-chat-turns]");
          if (turnsEl) turnsEl.innerHTML = historyView(workspace);
          const titleWrap = document.querySelector("[data-thread-title-wrap]");
          if (titleWrap && freshSession?.title) {
            titleWrap.classList.remove("muted");
            titleWrap.innerHTML = `${icon("message")}<span class="thread-title-text" title="${e(freshSession.title)}">${e(freshSession.title)}</span>`;
          }
          if (freshSession?.turns?.length) leaveEmptyState();
          const artifacts = studioArtifacts(freshSession);
          if (state.contextOpen && state.studioTab === "deck" && artifacts.length) {
            await refreshCanvasLive(workspace, { selectId: artifacts[0].id, flash: false });
          }
          syncActiveChatView({ skipRevalidate: true });
        }
      }).catch(() => {});
    }

    // 后台补齐或校准关联任务与项目
    if (bound && (!task || taskFromCache)) {
      void api("/assignments/" + encodeURIComponent(bound)).then((freshTask) => {
        if (state.me?.id !== userId) return;
        taskCache.set(bound, freshTask);
        if (state.chatWorkspace !== workspace) return;
        workspace.task = freshTask;
        updateContextDom(workspace);
        syncChatSubmit(chatForm(), workspace);
      }).catch(() => {});
    }
    if (group && (!project || projectFromCache)) {
      void api("/projects/" + encodeURIComponent(group)).then((freshProject) => {
        if (state.me?.id !== userId) return;
        projectCache.set(group, freshProject);
        if (state.chatWorkspace !== workspace) return;
        workspace.project = freshProject;
        updateContextDom(workspace);
      }).catch(() => {});
    }

    // 后台刷新侧栏会话列表与今日待办（不阻塞当前页切换）
    void api("/conversations").then((freshList) => {
      if (state.me?.id !== userId || !Array.isArray(freshList)) return;
      state.conversations = freshList;
      if (state.chatWorkspace === workspace) workspace.conversations = freshList;
      emitConversationsUpdated();
    }).catch(() => {});

    void api("/me/today").then((freshToday) => {
      if (state.me?.id !== userId || !freshToday?.items) return;
      state.today = freshToday.items;
      if (freshToday.today && state.me) state.me.today = freshToday.today;
      if (state.chatWorkspace === workspace) {
        workspace.tasks = freshToday.items.filter((item) => active(item.status));
        if (!workspace.task && !workspace.project) updateContextDom(workspace);
      }
    }).catch(() => {});
  };

  return workspace;
}

export function conversationNav(workspace) {
  if (!workspace) return "";
  return `<div class="workspace-nav">
    <button type="button" class="button new-session" data-act="route" data-id="chat" title="新建会话">
      ${icon("plus")}<span class="new-session-label">新建会话</span>
    </button>
    ${workspace.conversations.length ? `<div class="nav-label-row">
      <p class="nav-label">最近会话</p>
      ${typeof window.officeDesktop?.generateTitle === "function"
        ? `<button type="button" class="nav-label-action" data-act="chat-auto-titles" title="让 AI 根据内容为最近会话生成精炼标题" aria-label="AI 智能命名最近会话">${icon("sparkles")}</button>`
        : ""}
    </div>
      <button type="button" class="button ghost history-trigger" data-act="chat-history" title="最近会话" aria-label="最近会话">${icon("history")}</button>
      <div class="recent-sessions" data-recent-sessions>${workspace.conversations.map((session) => {
        const selected = state.route[0] === "chat" && state.route[1] === session.id;
        const running = Boolean(activeTurnFor(session.id));
        const pinned = Boolean(session.pinned);
        return `<div class="session-row ${selected ? "active" : ""} ${pinned ? "is-pinned" : ""} ${running ? "is-running" : ""}" draggable="true" data-session-row data-session-id="${e(session.id)}" data-pinned="${pinned ? "1" : "0"}">
          <button type="button" class="session-link ${selected ? "active" : ""} ${running ? "is-running" : ""}" data-act="route" data-id="chat/${e(session.id)}"
            title="${e(session.title)}${pinned ? "（已置顶）" : ""}${running ? "（正在执行）" : ""} · 可拖拽排序或拖入下方删除，右键更多操作" ${selected ? 'aria-current="page"' : ""}>
            ${icon(pinned ? "pin" : "message")}<span>${e(session.title)}</span>${running ? '<i class="session-running-dot" aria-label="正在执行"></i>' : ""}
          </button>
          <div class="session-row-actions">
            <button type="button" class="session-act-btn" data-act="session-menu" data-id="${e(session.id)}" title="置顶、重命名或更多" aria-label="会话操作菜单">${icon("more")}</button>
            <button type="button" class="session-act-btn is-danger" data-act="session-delete" data-id="${e(session.id)}" title="删除会话" aria-label="删除会话">${icon("trash")}</button>
          </div>
        </div>`;
      }).join("")}</div>
      <div class="session-trash-zone" data-session-trash-zone aria-hidden="true">
        ${icon("trash")}<span data-session-trash-label>拖到此处删除会话</span>
      </div>` : ""}
  </div>`;
}

export function showConversationHistory() {
  const sessions = state.chatWorkspace?.conversations || state.conversations || [];
  modal("最近会话", `<div class="stack-sm">${sessions.map((session) =>
    `<div class="session-history-row">
      <button type="button" class="button session-history-open" data-act="route" data-id="chat/${e(session.id)}">${icon(session.pinned ? "pin" : "message")}<span>${e(session.title)}</span></button>
      <button type="button" class="icon-button sm" data-act="session-pin" data-id="${e(session.id)}" title="${session.pinned ? "取消置顶" : "置顶"}">${icon("pin")}</button>
      <button type="button" class="icon-button sm" data-act="session-rename" data-id="${e(session.id)}" title="重命名">${icon("edit")}</button>
      <button type="button" class="icon-button sm" data-act="session-delete" data-id="${e(session.id)}" title="删除会话">${icon("trash")}</button>
    </div>`
  ).join("") || '<p class="hint">暂无历史会话</p>'}</div>`);
}

export async function copyReply(index, triggerButton = null) {
  const turn = state.chatWorkspace?.session?.turns?.[Number(index)];
  const text = turn?.reply;
  if (!text) return;
  const ok = await copyText(text, "已复制回复内容");
  if (ok && triggerButton?.isConnected) {
    const previous = triggerButton.innerHTML;
    triggerButton.classList.add("copied");
    triggerButton.innerHTML = `${icon("checkPlain")}<span>已复制</span>`;
    setTimeout(() => {
      if (triggerButton.isConnected) {
        triggerButton.classList.remove("copied");
        triggerButton.innerHTML = previous;
      }
    }, 1600);
  }
}

function artifactCard(file) {
  const size = file.size === undefined ? "" : file.size > 1048576 ? `${(file.size / 1048576).toFixed(1)} MB` : `${Math.ceil(file.size / 1024)} KB`;
  if (["pptx", "docx"].includes(file.extension)) {
    const isPpt = file.extension === "pptx";
    return `<div class="artifact-card" data-act="deck-open" data-id="${e(file.id)}" role="button" tabindex="0">
      <div class="artifact-card-main">
        <span class="artifact-badge ${e(file.extension)}">${icon(isPpt ? "pptx" : "docx")}</span>
        <span class="artifact-meta">
          <strong>${e(file.name)}</strong>
          <small>${isPpt ? "PowerPoint 演示文稿" : "Word 文稿"}${size ? " · " + size : ""}${file.parent_id ? " · 已保留旧版" : ""}</small>
        </span>
      </div>
      <div class="artifact-card-actions">
        <span class="artifact-open-hint">${icon("panelRight")}画布预览</span>
        <a class="text-button" href="/api/v1/files/${encodeURIComponent(file.id)}/download" download data-stop-prop>下载</a>
        <button class="text-button" type="button" data-act="edit-file" data-id="${e(file.id)}" data-stop-prop>改字</button>
      </div>
    </div>`;
  }
  return `<div class="session-file">${fileMarkup(file, { edit: true })}</div>`;
}

function turnMaterialsMarkup(materials) {
  const list = Array.isArray(materials) ? materials.filter((f) => f && f.name) : [];
  if (!list.length) return "";
  return `<div class="request-materials" aria-label="本轮附带的参考资料">${list.map((file) => {
    const ext = String(file.extension || "").toLowerCase();
    const isImg = imageExts.has(ext);
    const sizeText = formatFileSizeCompact(file.size);
    const act = file.id ? (isImg ? "preview-image" : "preview") : "";
    if (act) {
      return `<button type="button" class="request-material-chip" data-act="${act}" data-id="${e(file.id)}" data-name="${e(file.name)}" title="点击预览 ${e(file.name)}${sizeText ? `（${sizeText}）` : ""}">
        ${icon(sourceIconName(ext))}<span class="request-material-name">${e(file.name)}</span>
      </button>`;
    }
    return `<span class="request-material-chip">${icon(sourceIconName(ext))}<span class="request-material-name">${e(file.name)}</span></span>`;
  }).join("")}</div>`;
}

function liveTurnMarkup(live) {
  const selectionHtml = live.selection
    ? `<small class="request-scope">${e(selectionSummaryText(live.selection))}</small><br>`
    : "";
  const materialsHtml = turnMaterialsMarkup(live.materials);
  const steps = Array.isArray(live.steps) ? live.steps : [];
  const stepsHtml = steps.map((step, idx) =>
    `<li class="${idx === steps.length - 1 ? "active" : "done"}"><span class="step-dot" aria-hidden="true"></span><span>${e(step)}</span></li>`
  ).join("");
  const thinkingText = typeof live.thinking === "string" ? live.thinking : "";
  const hasThinking = Boolean(thinkingText.trim());
  const canStop = live.canCancel && live.phase !== "complete";
  return `<section class="conversation-turn pending-turn" data-live-thread="${e(live.threadId)}">
    <div class="request-entry"><div class="request-bubble">${materialsHtml}<p>${selectionHtml}${e(live.message)}</p></div></div>
    <div class="response-entry">
      <div class="response-head">
        <p class="response-byline"><span class="response-mark is-working">${icon("mark")}</span><strong>穹驭</strong></p>
      </div>
      <ol class="live-steps" data-progress-trace ${steps.length ? "" : "hidden"}>${stepsHtml}</ol>
      <details class="live-thinking" data-live-thinking ${hasThinking ? "open" : "hidden"}>
        <summary><span class="thinking-label">思考过程</span><span class="thinking-meta" data-thinking-meta>${hasThinking ? `${thinkingText.length} 字` : ""}</span></summary>
        <pre class="live-thinking-body" data-thinking-body>${e(thinkingText)}</pre>
      </details>
      <div class="running-bar">
        <span class="running-pulse" aria-hidden="true"></span>
        <p class="running-note" ${live.kind === "image" ? "data-image-progress" : ""} role="status">${e(live.statusText || "正在思考…")}</p>
        ${canStop ? `<button type="button" class="button stop-button" data-act="chat-cancel" ${live.cancelling ? "disabled" : ""}>${live.cancelling ? "正在停止…" : `${icon("stop")}停止`}</button>` : ""}
      </div>
    </div>
  </section>`;
}

function turnView(turn, index, canRetry = false) {
  const isOrphanRunning = turn.status === "running" && !turn.reply && !turn.error;
  const reply = turn.reply || turn.error || (turn.status === "running" ? "正在执行…" : "");
  const images = Array.isArray(turn.images) ? turn.images : [];
  const materialsHtml = turnMaterialsMarkup(turn.materials);
  const pipeline = turn.pipeline;
  const pipelineLabel = pipeline?.verdict === "pass" ? "验收通过" : pipeline?.verdict === "fail" ? "验收未通过" :
    pipeline?.verdict === "skipped" && pipeline?.tier === "light" ? "轻量任务 · 未验收" : "";
  const reviewLabels = { pass: "复核：通过", fail: "复核：未通过", not_needed: "复核：无需复核", unavailable: "复核：不可用" };
  return `<section class="conversation-turn">
    <div class="request-entry"><div class="request-bubble">${materialsHtml}<p>${turn.selection ? `<small class="request-scope">${e(selectionSummaryText(turn.selection))}</small><br>` : ""}${e(turn.message)}</p></div></div>
    ${reply || images.length ? `<div class="response-entry">
      <div class="response-head">
        <p class="response-byline"><span class="response-mark">${icon("mark")}</span><strong>穹驭</strong></p>
        <div class="response-tools">
          ${canRetry && turn.message ? `<button type="button" class="text-button turn-action-btn" data-act="chat-retry" data-id="${index}" title="带入输入框修改或重发">${icon("refresh")}<span>重用</span></button>` : ""}
          ${turn.reply && !turn.error ? `<button type="button" class="text-button copy-button" data-act="copy-reply" data-id="${index}" aria-label="复制回复">${icon("copy")}<span>复制</span></button>` : ""}
        </div>
      </div>
      ${turn.steps?.length ? `<details class="tool-trace"><summary>执行轨迹 · ${turn.steps.length} 步</summary>
        <ul>${turn.steps.map((step) => `<li>${icon("check")}<span>${e(step)}</span></li>`).join("")}</ul></details>` : ""}
      ${isOrphanRunning && clientExecutionRequired() && turn.id
        ? `<div class="running-bar orphan-running-bar">
            <p class="running-note" role="status">上一次桌面执行未正常结束，当前会话被锁定。</p>
            <button type="button" class="button stop-button" data-act="chat-abort-orphan" data-id="${e(turn.id)}" data-kind="${e(turn.kind === "image" ? "image" : "turn")}">${icon("stop")}结束并解锁</button>
          </div>`
        : turn.error ? `<p class="error">${e(reply)}</p>${feedbackErrorStrip(reply, { source: "chat", conversationId: state.chatWorkspace?.session?.id || "" })}` : `<div class="response-copy md-body">${renderMarkdown(reply)}</div>`}
      ${pipeline && pipelineLabel ? `<div class="pipeline-result">
        <div class="pipeline-result-status"><span class="pipeline-badge ${pipeline.verdict === "fail" ? "failed" : pipeline.verdict === "skipped" ? "skipped" : "passed"}">${e(pipelineLabel)}</span>
          <span class="pipeline-review">${e(reviewLabels[pipeline.review] || "复核：未知")}</span></div>
        ${pipeline.verdict === "fail" && Array.isArray(pipeline.issues) && pipeline.issues.length ?
          `<ul class="pipeline-issues">${pipeline.issues.map((issue) => `<li>${e(issue)}</li>`).join("")}</ul>` : ""}
      </div>` : ""}
      ${outputWarnings(turn.output_telemetry)}
      ${turn.file ? artifactCard(turn.file) : ""}
      ${images.length ? `<div class="turn-images">${images.map((image) => {
        const file = image.file || {};
        const id = typeof file.id === "string" ? file.id : "";
        const alt = typeof image.prompt === "string" ? image.prompt : "生成的图片";
        const quality = image.quality === "low" ? "草稿" : image.quality === "high" ? "精细" : "标准";
        const meta = `${image.width}×${image.height} · ${quality}质量 · ${image.model}`;
        return `<figure class="turn-image">
          <button type="button" class="turn-image-frame" data-act="image-lightbox" data-id="${e(id)}" data-meta="${e(meta)}" data-alt="${e(alt)}" aria-label="放大查看图片">
            <img src="/api/v1/files/${encodeURIComponent(id)}/image" alt="${e(alt)}" loading="lazy">
            <span class="turn-image-zoom">${icon("maximize")}</span>
          </button>
          <figcaption><span>${e(meta)}</span>
            <a href="/api/v1/files/${encodeURIComponent(id)}/download" download>下载</a></figcaption>
        </figure>`;
      }).join("")}</div>` : ""}
    </div>` : ""}
  </section>`;
}

function checkpointCardMarkup(workspace) {
  const cp = workspace?.checkpoint;
  const sessionId = workspace?.session?.id;
  if (!cp || !sessionId || cp.conversationId !== sessionId) return "";
  if (activeTurnFor(sessionId)) return "";
  if (workspace.task && !active(workspace.task.status)) return "";
  if (!clientExecutionRequired() || typeof window.officeDesktop?.runAgent !== "function") return "";
  const preserved = Array.isArray(cp.preserved) ? cp.preserved : [];
  const canResume = selectedModelAvailable(workspace) || Boolean(cp.hasCompletedResult);
  const errorText = cp.error ? cleanErrorMessage(cp.error, "") : "";
  const title = cp.hasCompletedResult
    ? "上次结果已生成，保存时中断"
    : `上次任务在「${cp.stageLabel || "任务执行"}」阶段中断`;
  return `<aside class="checkpoint-card" data-checkpoint-card role="region" aria-label="可续跑任务断点">
    <div class="checkpoint-card-head">
      <span class="checkpoint-badge">${icon("history")}<span>断点已保存</span></span>
      <strong class="checkpoint-title">${e(title)}</strong>
    </div>
    ${errorText ? `<p class="checkpoint-error">中断原因：${e(errorText)}</p>` : ""}
    ${preserved.length ? `<div class="checkpoint-chips">${preserved.map((item) => `<span class="checkpoint-chip">${icon("checkPlain")}<span>${e(item)}</span></span>`).join("")}</div>` : ""}
    <div class="checkpoint-actions">
      <button type="button" class="button primary sm checkpoint-resume-btn" data-act="chat-resume" ${canResume ? "" : "disabled"}>${icon("play")}${cp.hasCompletedResult ? "直接重新保存" : "从断点续跑"}</button>
      <button type="button" class="button sm" data-act="feedback-quick-upload" data-source="checkpoint" data-conversation-id="${e(sessionId)}" data-error="${e(errorText || title)}" title="一键打包并上传日志与断点状态">${icon("upload")}<span>一键上传日志</span></button>
      <button type="button" class="text-button" data-act="chat-discard-checkpoint">放弃断点</button>
    </div>
  </aside>`;
}

function historyView(workspace) {
  const session = workspace.session;
  const canRetry = !(workspace.task && !active(workspace.task.status));
  const activeLive = activeTurnFor(session?.id);
  const turns = session?.turns || [];
  let skippedRunning = false;
  const turnsHtml = turns.map((turn, i) => {
    if (activeLive && turn.status === "running" && !skippedRunning && (!activeLive.runId || turn.id === activeLive.runId)) {
      skippedRunning = true;
      return "";
    }
    return turnView(turn, i, canRetry);
  }).join("");
  const cpHtml = activeLive ? "" : checkpointCardMarkup(workspace);
  return `${session?.has_earlier ? `<div class="actions">${button("加载更早的消息", "chat-earlier", { icon: "history", cls: "sm" })}</div>` : ""}${turnsHtml}${activeLive ? liveTurnMarkup(activeLive) : cpHtml}`;
}

function welcomeHero(workspace) {
  const { task, project } = workspace;
  const firstName = state.me?.name || "";
  const heading = task ? task.title : project ? project.name : `${firstName ? `${firstName}，` : ""}今天要处理什么？`;
  const tag = task ? "关联任务" : project ? "项目空间 · 自动注入预整理上下文" : "";
  const subcopy = project?.description ? `<p class="welcome-subcopy">${e(project.description)}</p>` : "";
  return `<div class="welcome-hero">
    <div class="welcome-emblem" aria-hidden="true">${icon("mark")}</div>
    ${tag ? `<span class="welcome-tag">${e(tag)}</span>` : ""}
    <h1>${e(heading)}</h1>
    ${subcopy}
  </div>`;
}

function welcomeBelow(workspace) {
  const { task, project } = workspace;
  const examples = task ? [
    ["梳理交付要求", "先梳理这项任务的交付要求，列出需要补充的资料。"],
    ["生成初稿", task.kind === "pptx" ? "按任务要求生成一份可编辑 PPT，保留待确认的信息。" : "按任务要求生成一份可编辑 Word 初稿，保留待确认的信息。"]
  ] : project ? [
    ["梳理项目全貌与核心结论", "基于当前项目的预整理总览与共享资料，梳理项目目标、关键结论与下一步行动清单。"],
    ["按项目口径起草 Word 方案", "结合项目上下文中的背景目标、会议决策与口径规范，起草一份可编辑的 Word 方案文稿。"],
    ["基于项目资料制作汇报 PPT", "结合项目共享资料中的核心数据与结论，制作一份结构清晰的可编辑汇报 PPT。"]
  ] : [
    ["整理会议纪要", "帮我整理会议纪要，按议题、结论和后续行动展开。"],
    ["起草 Word 文档", "帮我起草一份可编辑 Word 文档。"],
    ["制作汇报 PPT", "帮我制作一份可编辑的汇报 PPT。"]
  ];

  const enabledProjectItems = (project?.context || []).filter((c) => c.enabled !== false);
  const projectSummaryCard = project
    ? `<section class="welcome-tasks welcome-project-card" aria-label="当前项目预整理上下文">
        <div class="between">
          <span class="context-label">${icon("sparkles")}<span>已自动挂载项目预整理上下文（${enabledProjectItems.length} 项生效）</span></span>
          <div class="actions">
            <button type="button" class="text-button" data-act="project-quick-add" data-id="${e(project.id)}">${icon("plus")}上传文件 / 记决策</button>
            <button type="button" class="text-button" data-act="route" data-id="projects/${e(project.id)}">项目空间${icon("right")}</button>
          </div>
        </div>
        ${enabledProjectItems.length
          ? `<div class="welcome-project-chips">${enabledProjectItems.slice(0, 6).map((item) => {
              const cat = CATEGORY_META[item.category] || CATEGORY_META.doc;
              return `<span class="welcome-project-chip" title="${e(item.digest?.summary || item.content || item.title)}">
                <small class="project-cat-badge is-${e(cat.tone)}">${e(cat.tag)}</small>
                <strong>${e(item.title)}</strong>
              </span>`;
            }).join("")}</div>`
          : `<p class="hint">该项目暂未添加资料，点击右上角「上传文件 / 记决策」即可直接拖入文档并自动预整理。</p>`}
      </section>`
    : "";

  const quickTasks = !task && !project && workspace.tasks?.length
    ? `<section class="welcome-tasks" aria-label="今日待办任务">
        <div class="between">
          <span class="context-label">今日待办</span>
          <div class="actions">
            <button type="button" class="text-button" data-act="chat-bind-project">${icon("layers")}关联项目上下文</button>
            <button type="button" class="text-button" data-act="route" data-id="today">全部${icon("right")}</button>
          </div>
        </div>
        <div class="welcome-task-grid">${workspace.tasks.slice(0, 4).map((item) => `
          <button type="button" class="welcome-task-item" data-act="route" data-id="chat/task/${encodeURIComponent(item.id)}">
            <strong>${e(item.title)}</strong>
            <span class="${item.overdue ? "late" : ""}">${e(deadline(item.due_at))}${item.overdue ? " · 逾期" : ""}</span>
          </button>`).join("")}
        </div>
      </section>`
    : (!task && !project
      ? `<div class="welcome-bind-row">
          <button type="button" class="text-button" data-act="chat-bind-project">${icon("layers")}<span>为当前会话关联「项目知识库上下文」（自动注入项目背景、决策与共享文档）</span></button>
        </div>`
      : "");

  return `<div class="welcome-below">
    <div class="prompt-examples">${examples.map(([label, value]) =>
      button(label, "chat-example", { id: value })).join("")}</div>
    ${projectSummaryCard}
    ${quickTasks}
  </div>`;
}

function contextAnchorBar(workspace) {
  const { task, project } = workspace;
  if (!task && !project) return "";
  if (task) {
    return `<div class="context-anchor">
      <div class="context-anchor-main">
        <span class="context-anchor-tag">关联任务</span>
        <strong title="${e(task.title)}">${e(task.title)}</strong>
        ${badge(task.status)}
        <span class="hint mono">${e(deadline(task.due_at))}</span>
      </div>
      <div class="context-anchor-actions">
        <button type="button" class="text-button" data-act="toggle-context" data-id="context">任务要求</button>
        <button type="button" class="text-button" data-act="open-task" data-id="${e(task.id)}">${active(task.status) ? "去交付" : "详情"}${icon("right")}</button>
      </div>
    </div>`;
  }
  const allItems = Array.isArray(project.context) ? project.context : [];
  const enabledItems = allItems.filter((c) => c.enabled !== false);
  const fileCount = enabledItems.filter((c) => c.artifact_id).length;
  return `<div class="context-anchor is-project-anchor">
    <div class="context-anchor-main">
      <span class="context-anchor-tag">项目空间</span>
      <strong title="${e(project.name)}">${e(project.name)}</strong>
      <span class="context-anchor-status ${enabledItems.length ? "is-ready" : ""}">
        ${icon(enabledItems.length ? "checkPlain" : "sparkles")}
        <span>${enabledItems.length ? `已自动注入 ${enabledItems.length} 项预整理上下文${fileCount ? `（${fileCount} 份文档）` : ""}` : "待添加项目上下文"}</span>
      </span>
    </div>
    <div class="context-anchor-actions">
      <button type="button" class="text-button" data-act="project-quick-add" data-id="${e(project.id)}" title="直接上传文件或补充决策到该项目">${icon("plus")}添加上下文</button>
      <button type="button" class="text-button" data-act="toggle-context" data-id="context">预整理详情</button>
      <button type="button" class="text-button" data-act="chat-bind-project" title="切换或解绑关联项目">切换项目</button>
      <button type="button" class="text-button" data-act="route" data-id="projects/${e(project.id)}">项目页${icon("right")}</button>
    </div>
  </div>`;
}

function sessionHistoryMaterials(workspace) {
  const map = new Map();
  const turns = workspace?.session?.turns || [];
  for (const turn of turns) {
    for (const file of turn?.materials || []) {
      if (file?.id && file?.name && !map.has(file.id)) map.set(file.id, file);
    }
  }
  const activeLive = activeTurnFor(workspace?.session?.id);
  for (const file of activeLive?.materials || []) {
    if (file?.id && file?.name && !map.has(file.id)) map.set(file.id, file);
  }
  return [...map.values()];
}

function contextBody(workspace) {
  const task = workspace.task;
  const project = workspace.project;
  const draft = draftFor(workspace);
  const currentSelected = draft.selected || [];
  const currentIds = new Set(currentSelected.map((f) => f.id));
  const pending = pendingUploadListFor(workspace.draftKey);
  const historyMaterials = sessionHistoryMaterials(workspace);
  const previousOnly = historyMaterials.filter((f) => !currentIds.has(f.id));
  const hasAnyMaterials = currentSelected.length > 0 || pending.length > 0 || previousOnly.length > 0;

  const materialsSection = `<section class="context-group">
    <div class="between">
      <p class="context-label">本会话参考资料</p>
      <button type="button" class="text-button" data-act="chat-sources">管理资料</button>
    </div>
    ${!hasAnyMaterials ? '<p class="hint">拖拽文档到会话区，或点输入框左下角回形针/文件夹即可添加。</p>' : ""}
    ${(currentSelected.length || pending.length) ? `<div class="context-subblock">
      <small class="context-sublabel">当前已挂载（下轮将发给 AI）</small>
      <ul class="context-files context-source-list">
        ${currentSelected.map((file) => {
          const ext = String(file.extension || "").toLowerCase();
          const isImg = imageExts.has(ext);
          return `<li class="context-source-row">
            <button type="button" class="context-source-link" data-act="${isImg ? "preview-image" : "preview"}" data-id="${e(file.id)}" data-name="${e(file.name)}" title="点击预览 ${e(file.name)}">
              ${icon(sourceIconName(ext))}<span class="context-source-name">${e(file.name)}</span>
            </button>
            <span class="context-source-tag is-ok">${isImg ? "已存文件库" : "已上传"}</span>
            <button type="button" class="context-source-remove" data-act="chat-remove-source" data-id="${e(file.id)}" title="移除该文件" aria-label="移除 ${e(file.name)}">${icon("x")}</button>
          </li>`;
        }).join("")}
        ${pending.map((item) => `<li class="context-source-row">
          <span class="context-source-link">${icon(sourceIconName(item.extension))}<span class="context-source-name">${e(item.name)}</span></span>
          <span class="context-source-tag ${item.status === "error" ? "is-error" : "is-uploading"}">${item.status === "error" ? "上传失败" : "上传中…"}</span>
          <button type="button" class="context-source-remove" data-act="chat-cancel-upload" data-id="${e(item.tempId)}" title="${item.status === "error" ? "关闭" : "取消上传"}" aria-label="取消 ${e(item.name)}">${icon("x")}</button>
        </li>`).join("")}
      </ul>
    </div>` : ""}
    ${previousOnly.length ? `<div class="context-subblock">
      <small class="context-sublabel">本会话已读取过</small>
      <ul class="context-files context-source-list">
        ${previousOnly.map((file) => {
          const ext = String(file.extension || "").toLowerCase();
          return `<li class="context-source-row">
            <button type="button" class="context-source-link" data-act="preview" data-id="${e(file.id)}" data-name="${e(file.name)}" title="点击预览 ${e(file.name)}">
              ${icon(sourceIconName(ext))}<span class="context-source-name">${e(file.name)}</span>
            </button>
            ${readable.has(ext) ? `<button type="button" class="text-button context-source-reattach" data-act="chat-reattach-source" data-id="${e(file.id)}" title="重新挂载到当前输入框">再次引用</button>` : ""}
          </li>`;
        }).join("")}
      </ul>
    </div>` : ""}
  </section>`;

  const projectItems = Array.isArray(project?.context) ? project.context : [];
  const enabledProjectItems = projectItems.filter((c) => c.enabled !== false);
  const projectBrief = (project?.overview?.effective_brief || project?.brief || "").trim();
  const projectSection = project
    ? `<section class="context-group context-project-group">
        <div class="between">
          <p class="context-label">项目预整理上下文</p>
          <div class="actions">
            <button type="button" class="text-button" data-act="project-quick-add" data-id="${e(project.id)}" title="直接上传文件或补充决策到项目">${icon("plus")}传文件/记决策</button>
            <button type="button" class="text-button" data-act="chat-bind-project" title="切换或解绑项目">切换</button>
          </div>
        </div>
        <div class="context-project-head">
          <h3>${e(project.name)}</h3>
          <span class="context-source-tag is-ok">${icon("checkPlain")}<span>已注入 ${enabledProjectItems.length}/${projectItems.length} 项</span></span>
        </div>
        ${project.description ? `<p class="context-copy">${e(project.description)}</p>` : ""}
        ${projectBrief ? `<details class="context-project-brief">
          <summary>${icon("sparkles")}<span>查看项目预整理总览大纲</span></summary>
          <div class="context-project-brief-body md-body">${renderMarkdown(projectBrief)}</div>
        </details>` : ""}
        ${projectItems.length ? `<ul class="context-files context-source-list">${projectItems.slice(0, 12).map((item) => {
          const cat = CATEGORY_META[item.category] || CATEGORY_META.doc;
          const isOn = item.enabled !== false;
          const summaryText = item.digest?.summary || item.content || "";
          return `<li class="context-source-row context-project-item ${isOn ? "" : "is-muted"}">
            <div class="context-project-item-main">
              <div class="context-project-item-top">
                <span class="project-cat-badge is-${e(cat.tone)}">${e(cat.tag)}</span>
                ${item.pinned ? `<span class="project-pin-mini" title="置顶优先">${icon("pin")}</span>` : ""}
                ${item.artifact_id
                  ? `<button type="button" class="context-source-link" data-act="preview" data-id="${e(item.artifact_id)}" data-name="${e(item.filename || item.title)}" title="点击预览 ${e(item.filename || item.title)}">
                      <span class="context-source-name">${e(item.title)}</span>
                    </button>`
                  : `<span class="context-source-name"><strong>${e(item.title)}</strong></span>`}
                <button type="button" class="context-source-tag ${isOn ? "is-ok" : ""}" data-act="project-context-toggle" data-project-id="${e(project.id)}" data-id="${e(item.id)}" data-enabled="${isOn ? "1" : "0"}" title="${isOn ? "当前已自动注入 AI（点击暂停）" : "当前已暂停（点击重新纳入 AI）"}">
                  ${isOn ? "已生效" : "已暂停"}
                </button>
              </div>
              ${summaryText ? `<small class="context-project-item-digest">${e(summaryText.slice(0, 110))}</small>` : ""}
            </div>
          </li>`;
        }).join("")}</ul>` : '<p class="hint">该项目暂无共享资料，点击右上角「传文件/记决策」即可添加并自动预整理。</p>'}
        <div class="actions">
          ${button("添加资料 / 记决策", "project-quick-add", { id: project.id, icon: "plus", cls: "sm primary" })}
          ${button("打开项目空间", "route", { id: "projects/" + project.id, icon: "right", cls: "sm" })}
        </div>
      </section>`
    : (!task ? `<section class="context-group">
        <div class="between">
          <p class="context-label">项目知识库上下文</p>
          <button type="button" class="text-button" data-act="chat-bind-project">${icon("layers")}关联项目</button>
        </div>
        <p class="hint">尚未关联项目空间。关联后，该项目下预整理的背景目标、会议决策、统一口径与多份共享文档将自动对本会话生效。</p>
        ${button("关联项目上下文", "chat-bind-project", { icon: "layers", cls: "sm" })}
      </section>` : "");

  return `<div class="studio-context-body">
    ${materialsSection}
    ${projectSection}
    ${task ? `<section class="context-group">
      <p class="context-label">关联任务</p>
      <h3>${e(task.title)}</h3>
      <div class="context-status">${badge(task.status)}<span>${e(deadline(task.due_at))}</span></div>
      ${task.description ? `<p class="context-copy">${e(task.description)}</p>` : ""}
      ${task.deliverable ? `<h4>交付要求</h4><p class="context-copy">${e(task.deliverable)}</p>` : ""}
      ${button(active(task.status) ? "回到任务交付" : "查看任务详情", "open-task", { id: task.id, icon: "right", cls: "sm" })}
    </section>` : project ? "" : `<section class="context-group">
      <p class="context-label">今日待办任务</p>
      ${workspace.tasks.length ? `<ul class="task-context-list">${workspace.tasks.map((item) => `<li><button type="button"
        data-act="route" data-id="chat/task/${encodeURIComponent(item.id)}">
        <strong>${e(item.title)}</strong><span>${e(deadline(item.due_at))}</span></button></li>`).join("")}</ul>`
        : '<p class="hint">今日暂无待办理任务</p>'}
      ${button("查看今日任务", "route", { id: "today", icon: "right", cls: "sm" })}
    </section>`}
    ${task?.materials.length ? `<section class="context-group"><p class="context-label">任务资料</p><ul class="context-files">
      ${task.materials.map((file) => `<li>${icon("clip")}<span>${e(file.name)}</span></li>`).join("")}</ul></section>` : ""}
  </div>`;
}

function studioPane(workspace, activeTab) {
  return `<aside class="studio-pane" aria-label="AI 实时创作画布与上下文">
    <header class="studio-pane-head">
      <div class="studio-tabs" role="tablist">
        <button type="button" role="tab" class="studio-tab ${activeTab === "deck" ? "active" : ""}" data-act="studio-tab" data-id="deck" aria-selected="${activeTab === "deck"}">实时画布</button>
        <button type="button" role="tab" class="studio-tab ${activeTab === "context" ? "active" : ""}" data-act="studio-tab" data-id="context" aria-selected="${activeTab === "context"}">办公上下文</button>
      </div>
      <button type="button" class="icon-button" data-act="toggle-context" data-id="close" aria-label="收起右侧画布" title="收起右侧画布">${icon("x")}</button>
    </header>
    <div class="studio-pane-content">
      <div class="studio-view" data-studio-pane="deck" ${activeTab === "deck" ? "" : "hidden"}>${deckPanel(workspace)}</div>
      <div class="studio-view" data-studio-pane="context" ${activeTab === "context" ? "" : "hidden"}>${contextBody(workspace)}</div>
    </div>
  </aside>`;
}

function selectedFiles(selected, pending = []) {
  const readyHtml = (Array.isArray(selected) ? selected : []).map((file) => {
    const ext = String(file.extension || "").toLowerCase();
    const isImg = imageExts.has(ext);
    const sizeText = formatFileSizeCompact(file.size);
    const statusLabel = isImg ? "已上传 · 文件库" : "已上传";
    const tip = isImg
      ? `${file.name}${sizeText ? `（${sizeText}）` : ""} · 图片已上传至「我的文件」，点击预览`
      : `${file.name}${sizeText ? `（${sizeText}）` : ""} · 已上传成功，发送时 AI 将读取，点击预览`;
    return `<span class="attached-file is-ready ${isImg ? "is-image" : ""}" data-file-id="${e(file.id)}">
      <button type="button" class="attached-main" data-act="${isImg ? "preview-image" : "preview"}" data-id="${e(file.id)}" data-name="${e(file.name)}" title="${e(tip)}">
        <span class="attached-icon">${icon(sourceIconName(ext))}</span>
        <span class="attached-name">${e(file.name)}</span>
        ${sizeText ? `<span class="attached-size">${e(sizeText)}</span>` : ""}
        <span class="attached-status is-ok">${icon("checkPlain")}<span>${e(statusLabel)}</span></span>
      </button>
      <button type="button" class="attached-remove" data-act="chat-remove-source" data-id="${e(file.id)}" title="取消移除 ${e(file.name)}" aria-label="移除 ${e(file.name)}">${icon("x")}</button>
    </span>`;
  }).join("");

  const pendingHtml = (Array.isArray(pending) ? pending : []).map((item) => {
    const isErr = item.status === "error";
    const sizeText = formatFileSizeCompact(item.size);
    return `<span class="attached-file ${isErr ? "is-error" : "is-uploading"}" data-temp-id="${e(item.tempId)}" title="${e(isErr ? (item.error || "上传失败") : `正在上传 ${item.name}…`)}">
      <span class="attached-main">
        <span class="attached-icon">${isErr ? icon("alert") : '<span class="attached-spinner" aria-hidden="true"></span>'}</span>
        <span class="attached-name">${e(item.name)}</span>
        ${sizeText ? `<span class="attached-size">${e(sizeText)}</span>` : ""}
        <span class="attached-status ${isErr ? "is-error" : "is-uploading"}">${e(isErr ? "上传失败" : "上传中…")}</span>
      </span>
      <button type="button" class="attached-remove" data-act="chat-cancel-upload" data-id="${e(item.tempId)}" title="${isErr ? "关闭失败提示" : "取消上传"}" aria-label="${isErr ? "关闭" : "取消上传"} ${e(item.name)}">${icon("x")}</button>
    </span>`;
  }).join("");

  return readyHtml + pendingHtml;
}

function composer(workspace) {
  const draft = draftFor(workspace);
  const desktopRequired = clientExecutionRequired();
  const desktopAvailable = typeof window.officeDesktop?.runAgent === "function";
  const modelSelected = selectedModelAvailable(workspace);
  const blocked = desktopRequired && !desktopAvailable;
  const modelRequired = desktopRequired && desktopAvailable;
  const imageModels = workspace.image_models || [];
  const showImageControls = desktopRequired && desktopAvailable && imageModels.length > 0;
  const imageAutoOn = showImageControls && imageAutoPreference();
  const drawerOpen = showImageControls && Boolean(state.imageDrawerOpen);
  const selectedCount = (draft.selected || []).length;
  const pendingList = pendingUploadListFor(workspace.draftKey);
  const draftText = draft.message || "";
  if (workspace.task && !active(workspace.task.status)) {
    return `<div class="composer-wrap"><div class="chat-composer">${blocked ? '<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>' : ""}
      <p class="notice">关联任务已完成或正在审核，当前仅可查阅历史会话与文件。</p></div></div>`;
  }
  return `<div class="composer-wrap">
  <form class="chat-composer" data-form="conversation" aria-label="发送消息">
    <div class="slash-menu" data-slash-menu hidden role="listbox" aria-label="快捷指令"></div>
    ${blocked ? '<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>' : ""}
    <div class="guidance-queue" data-chat-guidance>${guidanceQueueMarkup(workspace.session?.id || "")}</div>
    <div class="attached-files" data-chat-selection>${selectionChip(draft.selection)}</div>
    <div class="attached-files" data-chat-selected>${selectedFiles(draft.selected || [], pendingList)}</div>
    <div class="composer-grow" data-composer-grow data-replicated-value="${e(draftText + "\n")}">
      <textarea required name="message" rows="1" maxlength="6000" aria-label="办公请求"
        placeholder="描述你想完成的工作，输入 / 唤起指令…" ${blocked ? "disabled" : ""}>${e(draftText)}</textarea>
    </div>
    <div class="composer-actions">
      <div class="composer-tools">
        <label class="composer-tool upload-zone" data-upload-zone title="添加附件（也可直接拖拽文件到页面）" aria-label="添加附件">
          ${icon("clip")}<input type="file" data-upload accept=".docx,.pptx,.xlsx,.pdf,.txt,.md,.csv,.png,.jpg,.jpeg,.webp" multiple>
        </label>
        <button type="button" class="composer-tool ${selectedCount ? "has-value" : ""}" data-act="chat-sources" data-sources-btn title="引用已有资料" aria-label="引用资料">
          ${icon("folder")}<span class="tool-badge" data-sources-count ${selectedCount ? "" : "hidden"}>${selectedCount}</span>
        </button>
        <button type="button" class="composer-tool" data-act="slash-trigger" title="快捷指令（/）" aria-label="快捷指令">
          ${icon("slash")}
        </button>
        ${showImageControls ? `<button type="button" class="composer-tool ${drawerOpen ? "active" : ""} ${imageAutoOn ? "dot-active" : ""}" data-act="toggle-image-drawer" data-image-toggle aria-expanded="${drawerOpen ? "true" : "false"}" title="图片生成设置" aria-label="图片生成设置">
          ${icon("image")}
        </button>` : ""}
        <span class="composer-sep" aria-hidden="true"></span>
        <div class="composer-picker" data-picker="mode">${modePickerInner(draft.mode || "auto", blocked)}</div>
      </div>
      <div class="composer-right">
        ${modelRequired ? `<div class="composer-picker" data-picker="model">${modelPickerInner(workspace, blocked)}</div>` : ""}
        <button class="composer-send ${draftText.trim() ? "ready" : ""}" type="submit" aria-label="发送消息" title="发送（Enter）" ${blocked || (modelRequired && !modelSelected) ? "disabled" : ""}>${icon("arrowUp")}</button>
      </div>
    </div>
    ${showImageControls ? `<div class="image-drawer ${drawerOpen ? "open" : ""}" data-image-drawer aria-label="图片生成设置">
      <div class="image-drawer-inner">
        <div class="composer-image">
          <div class="image-controls">
            <label><span>尺寸</span><select name="image_size" aria-label="生图尺寸">
              <option value="1024x1024">方形 1024×1024</option><option value="1536x1024">横版 1536×1024</option>
              <option value="1024x1536">竖版 1024×1536</option><option value="1536x864">16:9 1536×864</option>
            </select></label>
            <label><span>质量</span><select name="image_quality" aria-label="生图质量">
              <option value="low">草稿</option><option value="medium" selected>标准</option><option value="high">精细</option>
            </select></label>
            ${button("生成图片", "chat-generate-image", { cls: "sm", icon: "image", disabled: blocked })}
          </div>
          <label class="image-auto-option" data-image-auto-wrap><input type="checkbox" name="allow_ai_images" ${imageAutoPreference() ? "checked" : ""}>
            <span>允许 AI 按需自动生图（不限张数，按张计费）</span></label>
        </div>
      </div>
    </div>` : ""}
    ${modelRequired && (workspace.modelNotice || workspace.pipelineReason) ? `<p class="chat-model-status ${workspace.models?.length ? "" : "error"}" data-chat-model-status role="status">${e([workspace.modelNotice, workspace.pipelineReason].filter(Boolean).join(" "))}</p>` : ""}
  </form>
  </div>`;
}

export function workspaceView(workspace) {
  const { session, task, project } = workspace;
  const activeLive = activeTurnFor(session?.id);
  const hasTurns = Boolean(session?.turns?.length || activeLive);
  const title = session?.title || (task ? task.title : project ? project.name : "新会话");
  const hasArtifact = studioArtifacts(session).length > 0;

  // 智能默认分屏策略：有 PPT / Word / 图片成果时默认并排开启右侧实时画布，无成果时默认居中沉浸
  if (workspace.draftKey !== state._lastStudioKey) {
    state._lastStudioKey = workspace.draftKey;
    state.contextOpen = hasArtifact;
    state.studioTab = hasArtifact ? "deck" : "context";
  }
  const activeTab = state.studioTab === "deck" ? "deck" : "context";
  const open = Boolean(state.contextOpen);

  return `<div class="agent-workspace ${open ? "studio-open" : ""} ${open && activeTab === "deck" ? "studio-deck" : ""}" data-has-deck="${hasArtifact ? "1" : "0"}">
    <section class="agent-thread ${hasTurns ? "" : "is-empty"}" data-chat-dropzone aria-label="办公会话">
      <div class="drop-overlay" aria-hidden="true">
        <div class="drop-overlay-box">
          ${icon("upload")}
          <strong>松开鼠标，作为参考资料加入当前会话</strong>
          <span>支持 Word、PowerPoint、Excel、PDF、TXT、Markdown、CSV 及常见图片</span>
        </div>
      </div>
      <header class="thread-header">
        <div class="thread-title ${hasTurns ? "" : "muted"}" data-thread-title-wrap>
          ${hasTurns ? `${icon("message")}<span class="thread-title-text" title="${e(title)}">${e(title)}</span>` : ""}
        </div>
        <div class="thread-header-actions">
          ${state.engine?.engine === "demo" ? '<span class="thread-engine">本地演示</span>' :
            state.engine?.model ? `<span class="thread-engine mono">${e(state.engine.model)}</span>` : ""}
          <button class="button sm ghost studio-trigger ${open && activeTab === "deck" ? "active" : ""}" type="button" data-act="toggle-context" data-id="deck" aria-expanded="${open && activeTab === "deck" ? "true" : "false"}">${icon("pptx")}画布</button>
          <button class="button sm ghost studio-trigger ${open && activeTab === "context" ? "active" : ""}" type="button" data-act="toggle-context" data-id="context" aria-expanded="${open && activeTab === "context" ? "true" : "false"}">${icon("panelRight")}上下文</button>
        </div>
      </header>
      ${contextAnchorBar(workspace)}
      <div class="thread-scroll">
        ${hasTurns ? "" : welcomeHero(workspace)}
        <div class="thread-content" data-chat-turns>${historyView(workspace)}</div>
      </div>
      ${composer(workspace)}
      ${hasTurns ? "" : welcomeBelow(workspace)}
    </section>
    ${studioPane(workspace, activeTab)}
  </div>`;
}

function chatForm() {
  return document.querySelector("form[data-form=conversation]");
}

export function pulseComposer() {
  const form = chatForm();
  if (!form) return;
  form.classList.remove("pulsing");
  void form.offsetWidth;
  form.classList.add("pulsing");
}

export function syncComposerGrow(textarea) {
  if (!textarea) return;
  const wrap = textarea.closest("[data-composer-grow]");
  if (wrap) wrap.dataset.replicatedValue = textarea.value + "\n";
  const sendBtn = textarea.form?.querySelector(".composer-send");
  if (sendBtn) sendBtn.classList.toggle("ready", Boolean(textarea.value.trim()));
}

export function toggleImageDrawer(forceOpen) {
  const form = chatForm();
  const drawer = form?.querySelector("[data-image-drawer]");
  const toggleBtn = form?.querySelector("[data-image-toggle]");
  if (!drawer) return;
  const next = typeof forceOpen === "boolean" ? forceOpen : !drawer.classList.contains("open");
  state.imageDrawerOpen = next;
  drawer.classList.toggle("open", next);
  if (toggleBtn) {
    toggleBtn.classList.toggle("active", next);
    toggleBtn.setAttribute("aria-expanded", String(next));
  }
}

export function syncImageAutoIndicator() {
  const form = chatForm();
  const toggleBtn = form?.querySelector("[data-image-toggle]");
  const checkbox = form?.elements?.allow_ai_images;
  if (toggleBtn) {
    toggleBtn.classList.toggle("dot-active", Boolean(checkbox ? checkbox.checked : imageAutoPreference()));
  }
}

function slashCommands() {
  const workspace = state.chatWorkspace;
  const showImage = clientExecutionRequired() && typeof window.officeDesktop?.runAgent === "function" && (workspace?.image_models?.length > 0);
  const items = [
    { id: "ppt", cmd: "/ppt", label: "制作 PowerPoint 演示文稿", icon: "pptx" },
    { id: "word", cmd: "/word", label: "起草可编辑 Word 文档", icon: "docx" },
    { id: "auto", cmd: "/auto", label: "自动判断回复或文档格式", icon: "message" },
    ...(showImage ? [{ id: "image", cmd: "/image", label: "展开生图参数面板", icon: "image" }] : []),
    { id: "project", cmd: "/project", label: workspace?.project ? "上传项目文件或管理预整理上下文" : "关联项目知识库上下文", icon: "layers" },
    { id: "ref", cmd: "/ref", label: "引用已有办公文件作为资料", icon: "folder" },
    { id: "task", cmd: "/task", label: "打开右侧办公上下文与待办", icon: "panelRight" },
    { id: "clear", cmd: "/clear", label: "清空当前选区与已附资料", icon: "x" }
  ];
  return items;
}

export function closeSlashMenu() {
  const menu = document.querySelector("[data-slash-menu]");
  if (menu) menu.hidden = true;
}

export function updateSlashMenu(textarea) {
  const menu = textarea?.form?.querySelector("[data-slash-menu]");
  if (!menu) return;
  const val = textarea.value;
  if (!val.startsWith("/") || /\s/.test(val)) {
    menu.hidden = true;
    return;
  }
  const q = val.slice(1).toLowerCase();
  const matched = slashCommands().filter((item) => item.id.includes(q) || item.label.toLowerCase().includes(q));
  if (!matched.length) {
    menu.hidden = true;
    return;
  }
  menu.innerHTML = matched.map((item, idx) => `
    <button type="button" class="slash-item ${idx === 0 ? "active" : ""}" role="option" data-act="slash-pick" data-id="${e(item.id)}">
      ${icon(item.icon)}<strong class="mono">${e(item.cmd)}</strong><span>${e(item.label)}</span>
    </button>`).join("");
  menu.hidden = false;
}

export function triggerSlashMenu() {
  const form = chatForm();
  const textarea = form?.elements?.message;
  if (!textarea || textarea.disabled || form.dataset.busy) return;
  const menu = form.querySelector("[data-slash-menu]");
  if (menu && !menu.hidden && textarea.value === "/") {
    textarea.value = "";
    syncComposerGrow(textarea);
    menu.hidden = true;
    textarea.focus();
    return;
  }
  if (!textarea.value.startsWith("/")) {
    textarea.value = "/";
    syncComposerGrow(textarea);
  }
  textarea.focus();
  updateSlashMenu(textarea);
}

export function handleSlashKeydown(event) {
  const menu = event.target?.form?.querySelector("[data-slash-menu]");
  if (!menu || menu.hidden) return false;
  const items = [...menu.querySelectorAll("[data-act=slash-pick]")];
  if (!items.length) return false;
  const cur = items.findIndex((el) => el.classList.contains("active"));
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    event.preventDefault();
    const next = (cur + (event.key === "ArrowDown" ? 1 : -1) + items.length) % items.length;
    items.forEach((el, idx) => el.classList.toggle("active", idx === next));
    items[next].scrollIntoView({ block: "nearest" });
    return true;
  }
  if (event.key === "Enter" || event.key === "Tab") {
    event.preventDefault();
    const target = items[cur >= 0 ? cur : 0];
    if (target) void applySlashCommand(target.dataset.id);
    return true;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    closeSlashMenu();
    return true;
  }
  return false;
}

export async function applySlashCommand(id) {
  const form = chatForm();
  const textarea = form?.elements?.message;
  closeSlashMenu();
  if (!form || !textarea) return;
  if (textarea.value.startsWith("/")) {
    textarea.value = "";
    syncComposerGrow(textarea);
    saveWorkspaceInput(form);
  }
  if (id === "ppt" || id === "word" || id === "auto") {
    const modeMap = { ppt: "pptx", word: "docx", auto: "auto" };
    const nameMap = { ppt: "已切换为 PowerPoint 输出模式", word: "已切换为 Word 输出模式", auto: "已切换为自动判断模式" };
    form.elements.mode.value = modeMap[id];
    syncModePickerDom(form);
    saveWorkspaceInput(form);
    pulseComposer();
    toastMessage(nameMap[id], "info");
    textarea.focus();
  } else if (id === "image") {
    toggleImageDrawer(true);
    pulseComposer();
    textarea.focus();
  } else if (id === "project") {
    const ws = state.chatWorkspace;
    if (ws?.project?.id) await openQuickProjectContextModal(ws.project.id);
    else await openBindProjectModal("");
  } else if (id === "ref") {
    await showSources();
  } else if (id === "task") {
    switchStudioTab("context");
  } else if (id === "clear") {
    updateSources(state.chatWorkspace, () => []);
    const chip = form.querySelector("[data-chat-selection]");
    if (chip) chip.innerHTML = "";
    writeDraft(state.chatWorkspace, form, { selection: null, selected: [] });
    toastMessage("已清空当前选区与已附资料", "info");
    textarea.focus();
  }
}

export function openImageLightbox(id, alt = "", meta = "") {
  if (!id) return;
  modal("图片预览", `<div class="lightbox-wrap">
    <div class="lightbox-stage">
      <img src="/api/v1/files/${encodeURIComponent(id)}/image" alt="${e(alt || "生成的图片")}">
    </div>
    <div class="lightbox-foot">
      <span class="hint">${e(meta || alt)}</span>
      <a class="button primary sm" href="/api/v1/files/${encodeURIComponent(id)}/download" download>${icon("download")}下载原图</a>
    </div>
  </div>`);
}

export function retryTurn(index) {
  const turn = state.chatWorkspace?.session?.turns?.[Number(index)];
  const form = chatForm();
  if (!turn?.message || !form || form.dataset.busy) return;
  form.elements.message.value = turn.message;
  syncComposerGrow(form.elements.message);
  saveWorkspaceInput(form);
  pulseComposer();
  form.elements.message.focus();
}

export function saveWorkspaceInput(form) {
  if (form?.dataset?.busy) {
    syncChatSubmit(form, state.chatWorkspace);
    return;
  }
  writeDraft(state.chatWorkspace, form);
}

export function selectDesktopModel(select, workspace, statusSelector = "[data-chat-model-status]", draftKey = modelDraftKey) {
  const submit = select.form?.querySelector('button[type="submit"]');
  const allowed = select.value === PIPELINE_MODEL ? workspace?.pipelineAvailable === true : workspace?.models?.includes(select.value);
  if (!allowed) {
    if (workspace) workspace.selectedModel = "";
    if (submit) submit.disabled = true;
    return false;
  }
  workspace.selectedModel = select.value;
  if (submit) submit.disabled = Boolean(select.form?.dataset.busy);
  workspace.modelNotice = "";
  saveDraft(draftKey, { model: select.value });
  const status = document.querySelector(statusSelector);
  if (status) {
    status.textContent = workspace.pipelineReason || "";
    status.hidden = !workspace.pipelineReason;
  }
  syncImageAutoIndicator();
  return true;
}

export function selectChatModel(select) {
  return selectDesktopModel(select, state.chatWorkspace);
}

export function useExample(value) {
  const form = chatForm();
  if (!form || form.dataset.busy) return;
  form.elements.message.value = value;
  syncComposerGrow(form.elements.message);
  saveWorkspaceInput(form);
  pulseComposer();
  form.elements.message.focus();
}

function syncSourcesDom(workspace = state.chatWorkspace) {
  const activeWs = state.chatWorkspace;
  if (!workspace || !activeWs || workspace.draftKey !== activeWs.draftKey) return;
  const form = chatForm();
  const selected = draftFor(activeWs).selected || [];
  const pending = pendingUploadListFor(activeWs.draftKey);
  if (form) {
    const slot = form.querySelector("[data-chat-selected]");
    if (slot) slot.innerHTML = selectedFiles(selected, pending);
    const countEl = form.querySelector("[data-sources-count]");
    const btnEl = form.querySelector("[data-sources-btn]");
    if (countEl) {
      countEl.textContent = String(selected.length);
      countEl.hidden = selected.length === 0;
    }
    if (btnEl) btnEl.classList.toggle("has-value", selected.length > 0);
    syncChatSubmit(form, activeWs);
  }
  updateContextDom(activeWs);
}

function updateSources(workspace, change) {
  const targetWs = workspace || state.chatWorkspace;
  if (!targetWs) return;
  const form = (targetWs === state.chatWorkspace || targetWs.draftKey === state.chatWorkspace?.draftKey) ? chatForm() : null;
  const selected = change(draftFor(targetWs).selected || []);
  writeDraft(targetWs, form?.dataset?.busy ? null : form, { selected });
  syncSourcesDom(targetWs);
}

export function removeSource(id) {
  const workspace = state.chatWorkspace;
  if (!workspace || !id) return;
  const current = draftFor(workspace).selected || [];
  const removed = current.find((file) => file.id === id);
  updateSources(workspace, (sources) => sources.filter((file) => file.id !== id));
  if (removed) {
    const isBusy = Boolean(chatForm()?.dataset?.busy || activeTurnFor(currentLiveThreadId(workspace)));
    toastMessage(
      isBusy ? `已移除「${removed.name}」（后续轮次与重做将不再携带该文件）` : `已取消文件「${removed.name}」`,
      "info"
    );
  }
}

export function cancelPendingUpload(tempId) {
  const workspace = state.chatWorkspace;
  if (!workspace || !tempId) return;
  const key = workspace.draftKey;
  const list = pendingUploadListFor(key);
  const target = list.find((item) => item.tempId === tempId);
  if (target?.controller) {
    try { target.controller.abort(); } catch {}
  }
  const next = list.filter((item) => item.tempId !== tempId);
  if (next.length) pendingUploads.set(key, next);
  else pendingUploads.delete(key);
  syncSourcesDom(workspace);
  if (target && target.status === "uploading") {
    toastMessage(`已取消上传「${target.name}」`, "info");
  }
}

export function reattachSource(fileId) {
  const workspace = state.chatWorkspace;
  if (!workspace || !fileId) return;
  const historyList = sessionHistoryMaterials(workspace);
  const found = historyList.find((file) => file.id === fileId);
  if (!found) return;
  const current = draftFor(workspace).selected || [];
  if (current.some((f) => f.id === fileId)) {
    toastMessage(`「${found.name}」已在当前输入框中`, "info");
    return;
  }
  if (readableSourcesOnly(current).length >= 8) {
    toastMessage("每轮最多添加 8 份参考资料", "error");
    return;
  }
  updateSources(workspace, (sources) => [...sources, found]);
  toastMessage(`已重新挂载「${found.name}」`, "success");
}

export async function attachFiles(files, input) {
  const workspace = state.chatWorkspace;
  if (!workspace) throw new Error("请先打开会话后再添加附件");
  const rawList = Array.from(files || []);
  if (!rawList.length) return;
  const existing = draftFor(workspace).selected || [];
  const currentPending = pendingUploadListFor(workspace.draftKey);
  const existingDocCount = readableSourcesOnly(existing).length
    + currentPending.filter((p) => p.status === "uploading" && readable.has(p.extension)).length;

  const accepted = [];
  let addedDocCount = 0;
  for (const file of rawList) {
    const name = String(file?.name || "文件");
    const ext = name.includes(".") ? name.split(".").pop().toLowerCase() : "";
    if (!readable.has(ext) && !imageExts.has(ext)) {
      if (input) input.value = "";
      throw new Error(`「${name}」格式不支持，请选择 Word、PowerPoint、Excel、PDF、TXT、Markdown、CSV 或常见图片`);
    }
    if (readable.has(ext)) {
      if (existingDocCount + addedDocCount >= 8) {
        if (input) input.value = "";
        throw new Error("每轮最多添加 8 份参考资料");
      }
      addedDocCount += 1;
    }
    accepted.push({
      file,
      tempId: uuid(),
      name,
      extension: ext,
      size: typeof file.size === "number" ? file.size : 0,
      status: "uploading",
      error: "",
      controller: new AbortController()
    });
  }
  if (input) input.value = "";

  // 0 毫秒立刻在输入框上方渲染「上传中…」占位卡片与取消按钮，消除拖入无反应的焦虑
  const getKey = () => (state.chatWorkspace && (state.chatWorkspace === workspace || state.chatWorkspace.session?.id === workspace.session?.id))
    ? state.chatWorkspace.draftKey
    : workspace.draftKey;
  const appendPending = (items) => {
    const k = getKey();
    pendingUploads.set(k, [...pendingUploadListFor(k), ...items]);
    syncSourcesDom(state.chatWorkspace || workspace);
  };
  const removePendingById = (tempId) => {
    for (const [k, arr] of pendingUploads.entries()) {
      if (arr.some((it) => it.tempId === tempId)) {
        const next = arr.filter((it) => it.tempId !== tempId);
        if (next.length) pendingUploads.set(k, next);
        else pendingUploads.delete(k);
      }
    }
  };
  const markPendingError = (tempId, errMsg) => {
    for (const [k, arr] of pendingUploads.entries()) {
      let hit = false;
      const next = arr.map((it) => {
        if (it.tempId !== tempId) return it;
        hit = true;
        return { ...it, status: "error", error: errMsg, controller: null };
      });
      if (hit) pendingUploads.set(k, next);
    }
  };

  appendPending(accepted);

  // 第一时间把所有 File 读入内存，杜绝 Electron 拖放句柄失效
  const materializedPairs = [];
  for (const entry of accepted) {
    if (entry.controller.signal.aborted) continue;
    try {
      const readyFile = await materializeFile(entry.file);
      materializedPairs.push({ entry, readyFile });
    } catch (err) {
      const msg = cleanErrorMessage(err, `读取「${entry.name}」失败`);
      markPendingError(entry.tempId, msg);
      syncSourcesDom(state.chatWorkspace || workspace);
      recordClientLog("error", "upload", msg, { name: entry.name });
      toastMessage(msg, "error");
    }
  }

  let addedDocs = 0;
  let addedImgs = 0;
  for (const { entry, readyFile } of materializedPairs) {
    if (entry.controller.signal.aborted) {
      removePendingById(entry.tempId);
      syncSourcesDom(state.chatWorkspace || workspace);
      continue;
    }
    try {
      const uploaded = await uploadOneFile(readyFile, {
        assignment: (state.chatWorkspace || workspace).task?.id || "",
        signal: entry.controller.signal,
        preMaterialized: true
      });
      removePendingById(entry.tempId);
      const targetWs = (state.chatWorkspace && (state.chatWorkspace === workspace || state.chatWorkspace.session?.id === workspace.session?.id))
        ? state.chatWorkspace
        : workspace;
      updateSources(targetWs, (sources) => {
        if (sources.some((s) => s.id === uploaded.id)) return sources;
        return [...sources, uploaded];
      });
      const ext = String(uploaded.extension || entry.extension || "").toLowerCase();
      if (readable.has(ext)) addedDocs += 1;
      else if (imageExts.has(ext)) addedImgs += 1;
    } catch (err) {
      if (entry.controller.signal.aborted || err?.name === "AbortError") {
        removePendingById(entry.tempId);
        syncSourcesDom(state.chatWorkspace || workspace);
        continue;
      }
      const msg = cleanErrorMessage(err, `上传「${entry.name}」失败`);
      markPendingError(entry.tempId, msg);
      syncSourcesDom(state.chatWorkspace || workspace);
      recordClientLog("error", "upload", msg, { name: entry.name });
      toastMessage(msg, "error");
    }
  }

  if (addedDocs && addedImgs) {
    toastMessage(`已上传 ${addedDocs} 份参考资料与 ${addedImgs} 张图片`, "success");
  } else if (addedDocs) {
    toastMessage(addedDocs > 1 ? `已上传 ${addedDocs} 份参考资料` : "文件已上传并加入当前会话", "success");
  } else if (addedImgs) {
    toastMessage("图片已上传至「我的文件」（可点击预览或随时移除）", "success");
  }
}

export async function refreshChatProjectContext(projectId) {
  if (!projectId) return null;
  const fresh = await api("/projects/" + encodeURIComponent(projectId));
  projectCache.set(projectId, fresh);
  const ws = state.chatWorkspace;
  if (ws && ws.project?.id === projectId) {
    ws.project = fresh;
    updateContextDom(ws);
    const welcomeCard = document.querySelector(".welcome-project-card");
    if (welcomeCard && !ws.session?.turns?.length) {
      const below = document.querySelector(".welcome-below");
      if (below) below.outerHTML = welcomeBelow(ws);
    }
  }
  return fresh;
}

export async function bindConversationProject(projectId) {
  const ws = state.chatWorkspace;
  if (!ws) return;
  const nextPid = String(projectId || "").trim();
  if (ws.task) {
    throw new Error("关联任务的会话不能同时绑定项目");
  }
  if (ws.session?.id) {
    const updated = await api(`/conversations/${encodeURIComponent(ws.session.id)}`, {
      method: "PATCH",
      body: { project_id: nextPid }
    });
    ws.session.project_id = updated.project_id;
    const cached = sessionCache.get(ws.session.id);
    if (cached) cached.project_id = updated.project_id;
    if (Array.isArray(state.conversations)) {
      state.conversations = state.conversations.map((c) => (c.id === ws.session.id ? { ...c, project_id: updated.project_id } : c));
    }
  } else {
    history.replaceState(null, "", nextPid ? `#chat/project/${encodeURIComponent(nextPid)}` : "#chat");
    state.route = nextPid ? ["chat", "project", nextPid] : ["chat"];
  }
  if (nextPid) {
    const freshProj = await api("/projects/" + encodeURIComponent(nextPid));
    projectCache.set(nextPid, freshProj);
    ws.project = freshProj;
    state.contextOpen = true;
    state.studioTab = "context";
    syncStudioDom();
    toastMessage(`已关联项目「${freshProj.name}」，预整理上下文已自动注入`, "success");
  } else {
    ws.project = null;
    toastMessage("已解除当前会话的项目上下文绑定", "info");
  }
  updateContextDom(ws);
  if (!ws.session?.turns?.length) {
    const hero = document.querySelector(".welcome-hero");
    if (hero) hero.outerHTML = welcomeHero(ws);
    const below = document.querySelector(".welcome-below");
    if (below) below.outerHTML = welcomeBelow(ws);
  }
}

export async function showSources() {
  const workspace = state.chatWorkspace;
  if (!workspace) return;
  const personal = await api("/files");
  if (state.chatWorkspace !== workspace) return;
  const projectSharedFiles = (workspace.project?.context || [])
    .filter((item) => item.artifact_id && item.extension && readable.has(String(item.extension).toLowerCase()))
    .map((item) => ({
      id: item.artifact_id,
      name: item.filename || item.title,
      extension: item.extension,
      size: item.size
    }));
  const autoProjectFileIds = new Set(
    (workspace.project?.context || [])
      .filter((item) => item.enabled !== false && item.artifact_id)
      .map((item) => item.artifact_id)
  );
  const available = [
    ...(workspace.task?.materials || []),
    ...projectSharedFiles,
    ...personal.filter((file) => !file.assignment_id || file.assignment_id === workspace.task?.id)
  ].filter((file) => readable.has(file.extension));
  const selected = new Map((draftFor(workspace).selected || []).map((file) => [file.id, file]));
  for (const file of available) if (!selected.has(file.id)) selected.set(file.id, file);
  workspace.availableSources = selected;
  const chosen = new Set((draftFor(workspace).selected || []).map((file) => file.id));
  const enabledProjectCount = (workspace.project?.context || []).filter((c) => c.enabled !== false).length;
  const projectBanner = workspace.project
    ? `<div class="project-quick-banner">
        <span class="project-digest-tag">${icon("checkPlain")}<span>项目「${e(workspace.project.name)}」已自动注入 ${enabledProjectCount} 项预整理上下文（含 ${autoProjectFileIds.size} 份项目文档，无需重复勾选）</span></span>
        <button type="button" class="text-button" data-act="project-quick-add" data-id="${e(workspace.project.id)}">${icon("plus")}上传项目资料</button>
      </div>`
    : (!workspace.task
      ? `<div class="project-quick-banner">
          <span class="hint">需要携带整个项目的背景、会议决策与多份资料？</span>
          <button type="button" class="text-button" data-act="chat-bind-project">${icon("layers")}关联项目上下文</button>
        </div>`
      : "");
  modal("引用资料", `<form class="stack" data-form="chat-sources">
    ${projectBanner}
    <div class="file-list">${[...selected.values()].map((file) => fileMarkup(file, { check: "ids", checked: chosen.has(file.id) })).join("") ||
      '<p class="hint">当前可引用的文件为空。可以从输入区直接拖拽或添加附件。</p>'}</div>
    <div class="actions"><button class="button primary" type="submit">使用所选资料</button></div></form>`);
}

export function chooseSources(form) {
  const workspace = state.chatWorkspace;
  const ids = new FormData(form).getAll("ids");
  if (ids.length > 8) throw new Error("每轮最多选择 8 份资料");
  updateSources(workspace, () => ids.map((id) => workspace.availableSources.get(id)).filter(Boolean));
}

export function scrollToLatest() {
  const container = document.querySelector(".thread-scroll");
  if (container) container.scrollTop = container.scrollHeight;
}

export async function loadEarlier() {
  const workspace = state.chatWorkspace;
  const session = workspace.session;
  if (!session?.has_earlier) return;
  const prior = await api("/conversations/" + encodeURIComponent(session.id) + "?offset=" + session.turns.length);
  if (state.chatWorkspace !== workspace) return;
  const container = document.querySelector(".thread-scroll");
  const height = container.scrollHeight;
  const position = container.scrollTop;
  session.turns = [...prior.turns, ...session.turns];
  session.has_earlier = prior.has_earlier;
  container.querySelector("[data-chat-turns]").innerHTML = historyView(workspace);
  container.scrollTop = position + container.scrollHeight - height;
}

async function readEvents(response, progress) {
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    const error = new Error(typeof body.detail === "string" ? body.detail : "发送失败，请检查连接后重试");
    error.retryAsNew = response.status < 500 && response.status !== 409;
    throw error;
  }
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffered = "", completed;
  function handle(line) {
    if (!line.trim()) return;
    const entry = JSON.parse(line);
    if (entry.type === "step") progress(entry.label);
    if (entry.type === "complete") completed = entry.turn;
    if (entry.type === "error") throw Object.assign(new Error(entry.message), { retryAsNew: true });
  }
  try {
    for (;;) {
      const { done, value } = await reader.read();
      buffered += decoder.decode(value, { stream: !done });
      const lines = buffered.split("\n");
      buffered = lines.pop();
      for (const line of lines) handle(line);
      if (done) { handle(buffered); break; }
    }
  } finally { reader.releaseLock(); }
  if (!completed) throw new Error("连接中断。可刷新会话确认执行结果。");
  if (completed.status === "failed") throw Object.assign(new Error(completed.error || "本轮未完成"), { retryAsNew: true });
  return completed;
}

function leaveEmptyState() {
  document.querySelector(".agent-thread")?.classList.remove("is-empty");
  document.querySelector(".welcome-hero")?.remove();
  document.querySelector(".welcome-below")?.remove();
}

function startProgress(form, live) {
  closeSlashMenu();
  leaveEmptyState();
  const messages = document.querySelector("[data-chat-turns]");
  if (messages) {
    findLiveProgress(live.threadId)?.remove();
    messages.insertAdjacentHTML("beforeend", liveTurnMarkup(live));
  }
  if (form?.elements?.message) {
    form.elements.message.value = "";
    syncComposerGrow(form.elements.message);
  }
  scrollToLatest();
  return findLiveProgress(live.threadId);
}

function setComposerLocked(form, locked, workspace = state.chatWorkspace) {
  if (!form?.isConnected) return;
  if (locked) closeComposerPickers();
  const threadId = currentLiveThreadId(workspace);
  const live = activeTurnFor(threadId);
  const allowGuidanceTyping = Boolean(locked && (!live || live.kind !== "image"));
  // 注意：思考过程中绝不锁死 chat-remove-source / chat-cancel-upload / chat-sources / input[type=file]，
  // 确保用户随时可以取消/移除已上传文件或管理附件。
  for (const field of form.querySelectorAll("select, button[data-act=slash-trigger], button[data-act=chat-generate-image], button[data-act=picker-toggle]")) {
    field.disabled = locked;
  }
  for (const removeBtn of form.querySelectorAll("button[data-act=chat-remove-source], button[data-act=chat-cancel-upload], button[data-act=chat-sources], input[type=file]")) {
    removeBtn.disabled = false;
  }
  const textarea = form.elements?.message;
  if (textarea) {
    textarea.disabled = locked && !allowGuidanceTyping;
    textarea.placeholder = allowGuidanceTyping
      ? "AI 正在思考，可随时输入补充要求或调整方向（Enter 追加引导）…"
      : "描述你想完成的工作，输入 / 唤起指令…";
  }
  if (!locked) {
    const modelSelect = form.querySelector("[data-chat-model]");
    if (modelSelect) modelSelect.disabled = !(workspace?.models?.length || workspace?.pipelineAvailable);
    const imgBtn = form.querySelector("[data-act=chat-generate-image]");
    if (imgBtn) {
      imgBtn.disabled = clientExecutionRequired() && typeof window.officeDesktop?.runAgent !== "function";
      imgBtn.innerHTML = `${icon("image")}生成图片`;
    }
    syncChatSubmit(form, workspace);
  } else {
    syncChatSubmit(form, workspace);
  }
  syncGuidanceDom(threadId);
}

function syncLiveCanvas(live) {
  if (!isViewingThread(live.threadId, live.userId)) return;
  if (live.liveKind || (state.contextOpen && state.studioTab === "deck")) {
    setCanvasLiveState({
      active: true,
      kind: live.liveKind || "pptx",
      label: live.statusText,
      prompt: live.message,
      steps: live.liveSteps,
      targetPage: live.targetPage
    });
  }
}

function recordLiveStep(live, label, { updateStatus = false } = {}) {
  if (typeof label !== "string" || !label) return;
  const isNewStep = live.steps[live.steps.length - 1] !== label;
  if (isNewStep) live.steps.push(label);
  if (live.liveSteps[live.liveSteps.length - 1] !== label) live.liveSteps.push(label);
  if (updateStatus) live.statusText = label;
  if (!isViewingThread(live.threadId, live.userId)) return;
  const node = findLiveProgress(live.threadId);
  if (node) {
    if (isNewStep) appendProgressStep(node, label);
    if (updateStatus) {
      const note = node.querySelector(".running-note");
      if (note) note.textContent = label;
    }
  }
  syncLiveCanvas(live);
}

function setLiveStatus(live, text, { removeCancel = false, canvasLabel = "" } = {}) {
  if (typeof text === "string" && text) live.statusText = text;
  if (canvasLabel && live.liveSteps[live.liveSteps.length - 1] !== canvasLabel) {
    live.liveSteps.push(canvasLabel);
  }
  if (!isViewingThread(live.threadId, live.userId)) return;
  const node = findLiveProgress(live.threadId);
  if (node) {
    const note = node.querySelector(".running-note");
    if (note && text) note.textContent = text;
    if (removeCancel) node.querySelector("[data-act=chat-cancel]")?.remove();
  }
  if (canvasLabel) syncLiveCanvas(live);
}

function ensureConversationListed(created, workspace) {
  if (!created?.id) return;
  if (!sessionCache.has(created.id)) {
    sessionCache.set(created.id, { ...created, turns: created.turns || [] });
  }
  const list = Array.isArray(state.conversations) ? state.conversations : [];
  if (!list.some((item) => item.id === created.id)) {
    state.conversations = [created, ...list];
    if (workspace && state.chatWorkspace === workspace) workspace.conversations = state.conversations;
  }
  emitConversationsUpdated();
}

export function syncActiveChatView({ skipRevalidate = false } = {}) {
  const workspace = state.chatWorkspace;
  if (!workspace || state.route[0] !== "chat") return;
  if (!skipRevalidate) workspace._revalidate?.();
  const threadId = workspace.session?.id || (state.route[1] && !["task", "project"].includes(state.route[1]) ? state.route[1] : "");
  const live = activeTurnFor(threadId);
  const form = chatForm();
  if (live) {
    if (form) {
      form.dataset.busy = "1";
      setComposerLocked(form, true, workspace);
      if (live.kind === "image") {
        const imgBtn = form.querySelector("[data-act=chat-generate-image]");
        if (imgBtn) {
          imgBtn.disabled = true;
          imgBtn.textContent = "正在生成图片…";
        }
      }
    }
    syncLiveCanvas(live);
    scrollToLatest();
  } else if (form && !form.dataset.busy) {
    syncChatSubmit(form, workspace);
  }
}

export async function cancelActiveChat() {
  const threadId = state.chatWorkspace?.session?.id ||
    (state.route[0] === "chat" && state.route[1] && !["task", "project"].includes(state.route[1]) ? state.route[1] : "");
  const live = activeTurnFor(threadId);
  if (!live || typeof live.cancel !== "function" || live.cancelling) return;
  live.cancelling = true;
  const cancel = live.cancel;
  const node = findLiveProgress(live.threadId);
  const stopBtn = node?.querySelector("[data-act=chat-cancel]") || document.querySelector("[data-act=chat-cancel]");
  if (stopBtn) {
    stopBtn.disabled = true;
    stopBtn.textContent = "正在停止…";
  }
  await cancel();
}

export async function abortOrphanTurn(runId, kind = "turn", triggerButton = null) {
  const workspace = state.chatWorkspace;
  const threadId = workspace?.session?.id;
  if (!workspace || !threadId || !runId) return;
  if (triggerButton?.isConnected) {
    triggerButton.disabled = true;
    triggerButton.textContent = "正在结束…";
  }
  const path = kind === "image"
    ? `/conversations/${encodeURIComponent(threadId)}/images/${encodeURIComponent(runId)}/fail`
    : `/conversations/${encodeURIComponent(threadId)}/turns/${encodeURIComponent(runId)}/fail`;
  try {
    await post(path, { message: "已手动结束未完成的任务，可重新发送" });
  } catch (error) {
    if (error.status !== 409 && error.status !== 404) {
      if (triggerButton?.isConnected) {
        triggerButton.disabled = false;
        triggerButton.innerHTML = `${icon("stop")}结束并解锁`;
      }
      throw error;
    }
  }
  sessionCache.delete(threadId);
  if (isViewingThread(threadId, state.me?.id)) {
    await refreshConversationInPlace(state.chatWorkspace, threadId);
    toastMessage("已解除锁定，可继续发送消息", "success");
  }
}

function appendProgressStep(progress, label) {
  if (!progress?.isConnected || typeof label !== "string" || !label) return;
  const list = progress.querySelector("[data-progress-trace]");
  if (!list) return;
  for (const prev of list.querySelectorAll("li.active")) {
    prev.classList.remove("active");
    prev.classList.add("done");
  }
  const item = document.createElement("li");
  item.className = "active";
  item.innerHTML = `<span class="step-dot" aria-hidden="true"></span><span>${e(label)}</span>`;
  list.append(item);
  list.hidden = false;
  scrollToLatest();
}

const THINKING_STAGE_LABELS = {
  understand: "理解需求",
  execute: "方案执行",
  accept: "独立验收",
  review: "高级复核",
  light: "轻量处理"
};

function appendLiveThinking(live, text, stage = "", model = "") {
  if (typeof text !== "string" || !text) return;
  const stageKey = `${stage || ""}|${model || ""}`;
  let chunk = text;
  if (stageKey && stageKey !== live.thinkingStage) {
    live.thinkingStage = stageKey;
    const title = [THINKING_STAGE_LABELS[stage] || "", model || ""].filter(Boolean).join(" · ");
    if (title) {
      chunk = `${live.thinking ? "\n\n" : ""}── ${title} ──\n${text}`;
    }
  }
  live.thinking = `${live.thinking || ""}${chunk}`;
  if (live.thinking.length > 24000) {
    live.thinking = "…（已省略更早思考）\n" + live.thinking.slice(-20000);
  }
  live.thinkingChars = (live.thinkingChars || 0) + text.length;
  live.statusText = `模型深度思考中（已思考 ${live.thinkingChars} 字）…`;
  if (!isViewingThread(live.threadId, live.userId)) return;
  const node = findLiveProgress(live.threadId);
  if (!node) return;
  const box = node.querySelector("[data-live-thinking]");
  const body = node.querySelector("[data-thinking-body]");
  const meta = node.querySelector("[data-thinking-meta]");
  const note = node.querySelector(".running-note");
  if (box && box.hidden) {
    box.hidden = false;
    box.open = true;
  }
  if (body) {
    body.textContent = live.thinking;
    body.scrollTop = body.scrollHeight;
  }
  if (meta) meta.textContent = `${live.thinkingChars} 字`;
  if (note) note.textContent = live.statusText;
}

function updateLiveProgress(live, chars, stage = "") {
  if (!Number.isFinite(chars) || chars <= 0) return;
  const label = THINKING_STAGE_LABELS[stage] || "方案";
  live.statusText = `正在生成${label}（已输出 ${chars} 字）…`;
  if (!isViewingThread(live.threadId, live.userId)) return;
  const node = findLiveProgress(live.threadId);
  const note = node?.querySelector(".running-note");
  if (note) note.textContent = live.statusText;
}

export async function runDesktopAgent(job, onStep = () => {}, afterRun = (result) => result, onEvent = () => {}) {
  const desktop = window.officeDesktop;
  if (typeof desktop?.runAgent !== "function") throw new Error("请在桌面客户端中使用 AI 办理");
  let unsubscribe = () => {};
  try {
    try {
      const stop = desktop.onAgentEvent?.((event) => {
        if (event?.type === "step") onStep(event.label);
        else if (event) onEvent(event);
      });
      if (typeof stop === "function") unsubscribe = stop;
    } catch {}
    const result = await desktop.runAgent(job);
    return await afterRun(result);
  } finally {
    try { unsubscribe(); } catch {}
  }
}

async function executeDesktopTurn(workspace, live, requestId, payload, { allowImageTool = true, resume = false } = {}) {
  const desktop = window.officeDesktop;
  if (typeof desktop?.runAgent !== "function") throw new Error("请在桌面客户端中使用 AI 办理");
  const model = workspace.selectedModel;
  if (!selectedModelAvailable(workspace)) throw new Error(workspace.models?.length ? "请先选择可用模型" : "暂无可用模型，请联系管理员开通。");

  const base = "/conversations/" + encodeURIComponent(live.threadId) + "/turns";
  let prepared;
  let cancelled = false;
  const prepareAbort = new AbortController();
  live.phase = "prepare";
  live.cancel = async () => {
    cancelled = true;
    if (live.phase === "prepare") prepareAbort.abort();
    else if (live.phase === "agent" && typeof desktop.cancelAgent === "function") {
      try { await desktop.cancelAgent(); } catch {}
    }
  };
  try {
    prepared = await api(base + "/prepare", { method: "POST", body: { id: requestId, ...payload }, signal: prepareAbort.signal });
    if (cancelled) throw new Error("已停止生成");
    if (!prepared?.run_id) throw new Error("服务端未返回本轮运行编号");
    live.runId = prepared.run_id;
    for (const step of prepared.steps || []) {
      recordLiveStep(live, step);
    }
    setLiveStatus(live, resume ? "正在从断点恢复执行…" : "正在调用模型…", { canvasLabel: resume ? "正在从断点恢复执行…" : "正在调用模型…" });
    live.phase = "agent";
    const enableImageTool = Boolean(workspace.image_models?.length) && allowImageTool === true;
    const job = {
      conversation_id: live.threadId, run_id: prepared.run_id, user_id: live.userId || "",
      system: prepared.system, prompt: prepared.prompt, context: prepared.context, model,
      resume: resume === true,
      checkpoint_meta: {
        message: payload.message,
        mode: payload.mode,
        material_ids: Array.isArray(payload.material_ids) ? payload.material_ids : [],
        ...(payload.selection ? { selection: payload.selection } : {})
      },
      ...(prepared.wants_pptx ? { ppt_check: true } : {}),
      ...(prepared.image_index_offset > 0 ? { image_index_offset: prepared.image_index_offset } : {}),
      ...(prepared.web_search ? { web_search: prepared.web_search } : {}),
      ...(enableImageTool ? { image_tool: true } : {})
    };
    return await runDesktopAgent(job, (label) => {
      if (typeof label !== "string" || !label) return;
      recordLiveStep(live, label, { updateStatus: true });
    }, async (result) => {
      const images = Array.isArray(result.images) ? result.images : [];
      let imageSaveFailures = 0;
      setLiveStatus(live, "正在保存结果…", { removeCancel: true, canvasLabel: "正在保存并同步画布…" });
      // 图片已在 sub2api 计费：即使用户已点停止，也先把它们存进本轮记录，避免已付费的图片丢失。
      for (const image of images) {
        try {
          await post(base + "/" + encodeURIComponent(prepared.run_id) + "/images", image);
        } catch {
          imageSaveFailures += 1;
        }
      }
      if (cancelled) {
        const saved = images.length - imageSaveFailures;
        throw new Error(images.length ? `已停止生成（本轮已生成并计费 ${images.length} 张图片，${saved} 张已保存在本轮记录中${imageSaveFailures ? `，${imageSaveFailures} 张保存失败` : ""}）` : "已停止生成");
      }
      live.phase = "complete";
      const completed = await post(base + "/" + encodeURIComponent(prepared.run_id) + "/complete", {
        text: result.text,
        usage: {
          ...(result.usage || {}), max_tokens_clamped: result.max_tokens_clamped,
          finish_reasons: result.finish_reasons, truncated: result.truncated
        },
        model: result.model || model,
        ...(result.pipeline ? { pipeline: result.pipeline } : {})
      });
      await desktop.clearCheckpoint?.(live.threadId).catch(() => {});
      if (imageSaveFailures) completed.image_save_failures = imageSaveFailures;
      return completed;
    }, (event) => {
      if (event.type === "thinking") appendLiveThinking(live, event.text, event.stage, event.model);
      else if (event.type === "progress") updateLiveProgress(live, event.chars, event.stage);
    });
  } catch (error) {
    if ((cancelled || error?.name === "AbortError") && !String(error?.message || "").startsWith("已停止生成")) {
      error = new Error("已停止生成");
    } else if (error instanceof Error) {
      error.message = cleanErrorMessage(error, "桌面执行失败");
    } else {
      error = new Error(cleanErrorMessage(error, "桌面执行失败"));
    }
    error.retryAsNew = true;
    if (prepared?.run_id) {
      try {
        await post(base + "/" + encodeURIComponent(prepared.run_id) + "/fail", {
          message: error.message || "桌面执行失败",
          ...(Array.isArray(live.steps) && live.steps.length ? { steps: live.steps.slice(-80) } : {})
        });
      } catch (failError) {
        if (live.phase === "complete" && failError.status === 409 && (!error.status || error.status >= 500)) {
          error.message = "无法确认本轮是否已保存，请刷新会话查看最新结果。";
        } else if (failError.status !== 409) error.message += `（失败状态同步未成功：${failError.message}）`;
      }
    }
    throw error;
  } finally {
    live.cancel = null;
  }
}

function inferLiveCanvasKind(message, mode, selection, workspace) {
  if (selection) return "pptx";
  if (mode === "pptx" || mode === "docx") return mode;
  if (mode === "reply") return "";
  const compact = String(message || "").toLowerCase();
  if (/pptx?|幻灯|演示文稿|汇报/.test(compact)) return "pptx";
  if (/\bdocx?\b|word|纪要|通知|报告|方案|文稿/.test(compact)) return "docx";
  if (state.contextOpen && state.studioTab === "deck" && studioArtifacts(workspace.session).length) {
    return studioArtifacts(workspace.session)[0].kind;
  }
  return "";
}

async function refreshConversationInPlace(workspace, threadId, { selectId = "", preferredPage = 1, openCanvas = false } = {}) {
  const [freshSession, conversations, freshCheckpoint] = await Promise.all([
    api("/conversations/" + encodeURIComponent(threadId)),
    api("/conversations").catch(() => null),
    loadCheckpointFor(threadId, state.me?.id)
  ]);
  sessionCache.set(threadId, freshSession);
  if (Array.isArray(conversations)) {
    state.conversations = conversations;
    emitConversationsUpdated();
  }
  const currentWorkspace = state.chatWorkspace;
  if (!currentWorkspace || (currentWorkspace !== workspace && currentWorkspace.session?.id !== threadId)) return;
  currentWorkspace.session = freshSession;
  currentWorkspace.checkpoint = freshCheckpoint;
  currentWorkspace.draftKey = "chat/" + threadId;
  if (Array.isArray(conversations)) currentWorkspace.conversations = conversations;
  const turnsEl = document.querySelector("[data-chat-turns]");
  if (turnsEl) turnsEl.innerHTML = historyView(currentWorkspace);
  const titleWrap = document.querySelector("[data-thread-title-wrap]");
  if (titleWrap && freshSession?.title) {
    titleWrap.classList.remove("muted");
    titleWrap.innerHTML = `${icon("message")}<span class="thread-title-text" title="${e(freshSession.title)}">${e(freshSession.title)}</span>`;
  }
  const selChip = document.querySelector("[data-chat-selection]");
  if (selChip) selChip.innerHTML = "";
  syncSourcesDom(currentWorkspace);
  updateContextDom(currentWorkspace);
  scrollToLatest();

  const artifacts = studioArtifacts(freshSession);
  if (openCanvas || selectId) {
    state.contextOpen = true;
    state.studioTab = "deck";
    state._lastStudioKey = "chat/" + threadId;
    syncStudioDom();
    await refreshCanvasLive(currentWorkspace, { selectId: selectId || artifacts[0]?.id || "", preferredPage, flash: true });
  } else {
    setCanvasLiveState({ active: false });
    if (state.contextOpen && state.studioTab === "deck") {
      await refreshCanvasLive(currentWorkspace, { selectId: artifacts[0]?.id || "", preferredPage, flash: false });
    }
  }
}

export async function generateImageFromComposer(triggerButton, rerender) {
  const workspace = state.chatWorkspace;
  const form = chatForm();
  if (!workspace || !form) throw new Error("当前会话不可用");
  if (form.dataset.busy) throw new Error("请等待当前请求执行结束");
  if (typeof window.officeDesktop?.generateImage !== "function") throw new Error("请在桌面客户端中生成图片");
  const prompt = form.elements.message.value.trim();
  if (!prompt) throw new Error("请先输入图片说明或提示词");
  const imageModels = workspace.image_models || [];
  const model = imageModels.includes("gpt-image-2.6-flare") ? "gpt-image-2.6-flare" : imageModels[0];
  if (!model) throw new Error("暂无可用的生图模型");
  const size = form.elements.image_size.value;
  const quality = form.elements.image_quality.value;
  const owner = state.me?.id;
  const threadId = workspace.session?.id || draftFor(workspace).newId || uuid();
  if (activeTurnFor(threadId)) throw new Error("当前会话有一轮任务正在执行");

  const live = {
    threadId,
    runId: null,
    userId: owner,
    kind: "image",
    message: prompt,
    selection: null,
    steps: [],
    statusText: "正在准备生图请求…",
    liveKind: "image",
    liveSteps: ["正在准备生图请求"],
    targetPage: null,
    canCancel: false,
    cancelling: false,
    phase: "prepare",
    cancel: null
  };
  activeTurns.set(threadId, live);
  emitConversationsUpdated();

  leaveEmptyState();
  switchStudioTab("deck");
  setCanvasLiveState({ active: true, kind: "image", label: "正在构图并绘制图片…", prompt, steps: live.liveSteps });
  startProgress(form, live);
  form.dataset.busy = "1";
  setComposerLocked(form, true, workspace);
  const originalLabel = triggerButton.textContent;
  triggerButton.disabled = true;
  triggerButton.textContent = "正在生成图片…";

  let prepared;
  let generated;
  let saved = false;
  try {
    if (!workspace.session) {
      const created = await post("/conversations", {
        id: threadId, assignment_id: workspace.task?.id || null,
        project_id: workspace.project?.id || null
      });
      workspace.session = created;
      ensureConversationListed(created, workspace);
      if (state.chatWorkspace === workspace && state.me?.id === owner) {
        workspace.draftKey = "chat/" + threadId;
        history.replaceState(null, "", "#chat/" + encodeURIComponent(threadId));
        state.route = ["chat", threadId];
      }
    }
    const base = "/conversations/" + encodeURIComponent(threadId) + "/images";
    prepared = await post(base + "/prepare", { prompt, model, size, quality });
    if (!prepared?.run_id) throw new Error("服务端未返回生图运行编号");
    live.runId = prepared.run_id;
    live.phase = "generate";
    setLiveStatus(live, "正在生成图片…", { canvasLabel: "正在渲染高分辨率画面…" });
    generated = await window.officeDesktop.generateImage({ prompt, model, size, quality });
    live.phase = "complete";
    setLiveStatus(live, "正在保存图片…", { canvasLabel: "正在保存并同步到画布…" });
    const completed = await post(base + "/" + encodeURIComponent(prepared.run_id) + "/complete", {
      png_base64: generated.png_base64, width: generated.width, height: generated.height,
      model: generated.model, size: generated.size, quality: generated.quality, usage: generated.usage || {}
    });
    saved = true;
    activeTurns.delete(threadId);
    sessionCache.delete(threadId);
    if (isViewingThread(threadId, owner)) {
      const currentForm = chatForm();
      if (currentForm?.elements?.message) {
        currentForm.elements.message.value = "";
        syncComposerGrow(currentForm.elements.message);
      }
      const latestImgId = completed?.images?.[0]?.file?.id || "";
      await refreshConversationInPlace(state.chatWorkspace, threadId, { selectId: latestImgId, openCanvas: true });
    } else if (state.me?.id === owner) {
      emitConversationsUpdated();
      void api("/conversations/" + encodeURIComponent(threadId)).then((fresh) => { sessionCache.set(threadId, fresh); }).catch(() => {});
      toastMessage("后台图片已生成，可切回原会话查看", "success");
    }
    void refreshDesktopUsage().catch(() => {});
    return completed;
  } catch (error) {
    activeTurns.delete(threadId);
    sessionCache.delete(threadId);
    emitConversationsUpdated();
    const message = prepared?.run_id && !saved && generated && live.phase === "complete"
      ? `图片已生成并计费，但保存失败：${error.message}`
      : (error.message || "生图失败");
    if (prepared?.run_id && !saved) {
      try {
        await post("/conversations/" + encodeURIComponent(threadId) + "/images/" + encodeURIComponent(prepared.run_id) + "/fail",
          { message: message.slice(0, 300) });
      } catch {}
    }
    if (isViewingThread(threadId, owner)) {
      setCanvasLiveState({ active: false });
      const node = findLiveProgress(threadId);
      if (node) {
        node.querySelector(".running-pulse")?.remove();
        node.querySelector(".response-mark")?.classList.remove("is-working");
        const note = node.querySelector(".running-note");
        if (note) { note.textContent = message; note.classList.add("error"); }
      }
      throw error;
    } else if (state.me?.id === owner) {
      toastMessage(`后台生图失败：${message}`, "error");
    }
  } finally {
    activeTurns.delete(threadId);
    emitConversationsUpdated();
    delete form.dataset.busy;
    if (form.isConnected) setComposerLocked(form, false, state.chatWorkspace);
    if (isViewingThread(threadId, owner)) {
      const currentForm = chatForm();
      if (currentForm && currentForm !== form) {
        delete currentForm.dataset.busy;
        setComposerLocked(currentForm, false, state.chatWorkspace);
      }
    }
    if (triggerButton.isConnected) {
      triggerButton.disabled = false;
      triggerButton.textContent = originalLabel;
    }
  }
}

async function runConversationTurn({
  workspace,
  form,
  message,
  payload,
  selected = [],
  draft = {},
  allowAiImages = true,
  resume = false,
  rethrowToForm = true
}) {
  const user = state.me.id;
  const signature = JSON.stringify(payload);
  const requestId = !resume && draft.pending?.signature === signature ? draft.pending.id : uuid();
  const threadId = workspace.session?.id || draft.newId || uuid();
  if (activeTurnFor(threadId)) {
    throw new Error("当前会话有一轮任务正在执行");
  }
  const other = otherActiveDesktopTurn(threadId);
  if (other) {
    const otherTitle = state.conversations?.find((c) => c.id === other.threadId)?.title;
    throw new Error(otherTitle
      ? `会话「${otherTitle}」正在本机执行，请等待完成或切换回去停止后再发送`
      : "另一会话正在本机执行任务，请等待完成或停止后再发送");
  }

  writeDraft(workspace, form, { newId: threadId, pending: { id: requestId, signature } });
  const liveKind = inferLiveCanvasKind(message, payload.mode, payload.selection || draft.selection, workspace);
  const canCancel = !clientExecutionRequired() || typeof window.officeDesktop?.cancelAgent === "function";
  const initialMaterials = readableSourcesOnly(selected);
  const live = {
    threadId,
    runId: null,
    userId: user,
    kind: "turn",
    message,
    selection: payload.selection || draft.selection || null,
    materials: initialMaterials,
    steps: [],
    statusText: resume ? "正在从断点恢复执行…" : "正在思考…",
    liveKind,
    liveSteps: [],
    targetPage: (payload.selection || draft.selection)?.slide || null,
    canCancel,
    cancelling: false,
    phase: "prepare",
    cancel: null
  };
  activeTurns.set(threadId, live);
  emitConversationsUpdated();
  updateContextDom(workspace);

  const initialInView = () => Boolean(form?.isConnected && state.me?.id === user);
  let pendingAutoNext = null;
  if (form) {
    form.dataset.busy = "1";
    setComposerLocked(form, true, workspace);
    form.querySelector(".error")?.remove();
  }
  document.querySelector("[data-checkpoint-card]")?.remove();
  try {
    if (!workspace.session) {
      const oldDraftKey = workspace.draftKey;
      const created = await post("/conversations", {
        id: threadId, assignment_id: workspace.task?.id || null,
        project_id: workspace.project?.id || null
      });
      workspace.session = created;
      ensureConversationListed(created, workspace);
      const nextDraftKey = "chat/" + threadId;
      const latestOldDraft = readDraft(oldDraftKey) || {};
      const migratedSelected = Array.isArray(latestOldDraft.selected) ? latestOldDraft.selected : selected;
      saveDraft(nextDraftKey, {
        ...(readDraft(nextDraftKey) || {}),
        selected: migratedSelected,
        mode: payload.mode
      });
      if (oldDraftKey && oldDraftKey !== nextDraftKey) {
        saveDraft(oldDraftKey, { message: "", selected: [], selection: null, newId: null, pending: null });
        if (pendingUploads.has(oldDraftKey)) {
          pendingUploads.set(nextDraftKey, [
            ...pendingUploadListFor(nextDraftKey),
            ...pendingUploadListFor(oldDraftKey)
          ]);
          pendingUploads.delete(oldDraftKey);
        }
      }
      workspace.draftKey = nextDraftKey;
      if (initialInView()) {
        history.replaceState(null, "", "#chat/" + encodeURIComponent(threadId));
        state.route = ["chat", threadId];
        updateContextDom(workspace);
      }
    }
    if (isViewingThread(threadId, user)) {
      startProgress(form?.isConnected ? form : chatForm(), live);
      if (liveKind) {
        switchStudioTab("deck");
        setLiveStatus(live, live.statusText, {
          canvasLabel: resume
            ? "正在从断点恢复并继续制作…"
            : ((payload.selection || draft.selection)?.slide ? `正在准备修改第 ${(payload.selection || draft.selection).slide} 页…` : "正在理解工作要求…")
        });
      }
    }
    let completed;
    if (clientExecutionRequired()) {
      completed = await executeDesktopTurn(workspace, live, requestId, payload, { allowImageTool: allowAiImages, resume });
    } else {
      const controller = new AbortController();
      live.cancel = async () => { controller.abort(); };
      try {
        const response = await fetch("/api/v1/conversations/" + encodeURIComponent(threadId) + "/turns", {
          method: "POST", headers: { "Content-Type": "application/json" }, credentials: "same-origin",
          signal: controller.signal,
          body: JSON.stringify({ id: requestId, ...payload })
        });
        completed = await readEvents(response, (label) => {
          recordLiveStep(live, label, { updateStatus: true });
        });
      } catch (error) {
        if (error?.name === "AbortError") throw Object.assign(new Error("已停止生成"), { retryAsNew: true });
        throw error;
      } finally {
        live.cancel = null;
      }
    }
    if (state.me?.id !== user) return;
    activeTurns.delete(threadId);
    sessionCache.delete(threadId);
    workspace.checkpoint = null;
    const latestThreadDraft = readDraft("chat/" + threadId) || readDraft(workspace.draftKey) || {};
    const currentSelected = Array.isArray(latestThreadDraft.selected) ? latestThreadDraft.selected : selected;
    writeDraft(workspace, null, { message: "", newId: null, pending: null, selection: null, selected: currentSelected });
    const latestFileId = completed.file && ["pptx", "docx"].includes(completed.file.extension)
      ? completed.file.id
      : (completed.images?.[0]?.file?.id || "");
    const targetPage = completed.file?.extension === "pptx" ? ((payload.selection || draft.selection)?.slide || 1) : 1;
    const deck = latestFileId ? { fileId: latestFileId, page: targetPage } : draft.deck;
    saveDraft("chat/" + threadId, { selected: currentSelected, mode: payload.mode, deck });
    if (completed.file && workspace.task) {
      const key = "submit/" + workspace.task.id;
      const submission = readDraft(key) || {};
      const files = new Set(submission.artifact_ids || []);
      if (completed.file.parent_id) files.delete(completed.file.parent_id);
      files.add(completed.file.id);
      saveDraft(key, { ...submission, artifact_ids: [...files], key: uuid() });
    }
    if (isViewingThread(threadId, user)) {
      await refreshConversationInPlace(state.chatWorkspace, threadId, {
        selectId: latestFileId,
        preferredPage: targetPage,
        openCanvas: Boolean(latestFileId)
      });
      if (completed.image_save_failures > 0) {
        const latest = document.querySelector("[data-chat-turns] .conversation-turn:last-child .response-entry");
        if (latest) {
          const warning = document.createElement("p");
          warning.className = "output-warnings";
          warning.textContent = `有 ${completed.image_save_failures} 张图片保存失败（已计费）`;
          latest.append(warning);
        }
      }
      chatForm()?.elements.message.focus({ preventScroll: true });
    } else {
      emitConversationsUpdated();
      void Promise.all([
        api("/conversations/" + encodeURIComponent(threadId)).then((fresh) => { sessionCache.set(threadId, fresh); }),
        api("/conversations").then((list) => {
          if (state.me?.id === user && Array.isArray(list)) {
            state.conversations = list;
            emitConversationsUpdated();
          }
        })
      ]).catch(() => {});
      const title = state.conversations?.find((c) => c.id === threadId)?.title;
      toastMessage(title ? `会话「${title}」已完成` : "后台会话已完成，可切回查看结果", "success");
    }
    void refreshDesktopUsage().catch(() => {});
    const curTitle = state.conversations?.find((c) => c.id === threadId)?.title || workspace.session?.title || "";
    if (!workspace.task && !workspace.project && looksLikeRawTitle(curTitle)) {
      void generateConversationTitle(threadId, { force: true }).catch(() => {});
    }
    const followups = guidanceListFor(threadId);
    if (followups.length && state.me?.id === user) {
      queuedGuidance.delete(threadId);
      syncGuidanceDom(threadId);
      const nextSel = Array.isArray((readDraft("chat/" + threadId) || {}).selected)
        ? (readDraft("chat/" + threadId) || {}).selected
        : currentSelected;
      pendingAutoNext = {
        kind: "followup",
        message: formatFollowUpGuidanceMessage(followups.map((item) => item.text)),
        mode: payload.mode,
        selected: nextSel,
        selection: null
      };
    }
  } catch (error) {
    activeTurns.delete(threadId);
    sessionCache.delete(threadId);
    emitConversationsUpdated();
    const key = workspace.session?.id ? "chat/" + threadId : workspace.draftKey;
    const latestDraftNow = readDraft(key) || readDraft(workspace.draftKey) || draft;
    const latestSelectedNow = Array.isArray(latestDraftNow.selected) ? latestDraftNow.selected : selected;
    const redirectTexts = Array.isArray(live.redirectGuidance) && live.redirectGuidance.length ? live.redirectGuidance : null;
    if (redirectTexts && state.me?.id === user) {
      if (isViewingThread(threadId, user)) {
        setCanvasLiveState({ active: false });
        findLiveProgress(threadId)?.remove();
      }
      pendingAutoNext = {
        kind: "redirect",
        message: formatMergedGuidanceMessage(message, redirectTexts),
        mode: payload.mode,
        selected: latestSelectedNow,
        selection: payload.selection || draft.selection || null
      };
    } else {
    const cleanedMessage = cleanErrorMessage(error, "本轮执行失败");
    recordClientLog("error", "chat", cleanedMessage, { conversationId: threadId, resume: Boolean(resume) });
    if (error instanceof Error) error.message = cleanedMessage;
    else error = new Error(cleanedMessage);
    if (state.me?.id === user) {
      const pending = error.retryAsNew ? null : { id: requestId, signature };
      saveDraft(key, { ...draft, selected: latestSelectedNow, message, mode: payload.mode, newId: threadId, pending });
      workspace.draftKey = key;
      workspace.checkpoint = await loadCheckpointFor(threadId, user);
    }
    if (isViewingThread(threadId, user)) {
      setCanvasLiveState({ active: false });
      const node = findLiveProgress(threadId);
      node?.querySelector("[data-act=chat-cancel]")?.remove();
      node?.querySelector(".running-pulse")?.remove();
      node?.querySelector(".response-mark")?.classList.remove("is-working");
      const current = node?.querySelector(".running-note");
      if (current) {
        current.textContent = cleanedMessage;
        current.classList.add("error");
        if (!workspace.checkpoint && !node.querySelector(".feedback-inline-bar")) {
          current.insertAdjacentHTML("afterend", feedbackErrorStrip(cleanedMessage, { source: "chat", conversationId: threadId }));
        }
      }
      if (workspace.checkpoint) syncCheckpointCardDom(workspace);
      const currentForm = chatForm();
      if (!workspace.checkpoint && currentForm?.elements?.message && !resume) {
        currentForm.elements.message.value = message;
        syncComposerGrow(currentForm.elements.message);
      }
      if (rethrowToForm && form?.isConnected && !workspace.checkpoint) throw error;
      if (!current && !workspace.checkpoint) toastMessage(cleanedMessage, "error");
    } else if (state.me?.id === user) {
      const stopped = cleanedMessage.startsWith("已停止生成");
      toastMessage(stopped ? "后台会话已停止" : `后台会话未完成（已保存断点）：${cleanedMessage}`, stopped ? "info" : "error");
    }
    }
  } finally {
    activeTurns.delete(threadId);
    emitConversationsUpdated();
    if (form?.isConnected) {
      delete form.dataset.busy;
      setComposerLocked(form, false, state.chatWorkspace);
    }
    if (isViewingThread(threadId, user)) {
      const currentForm = chatForm();
      if (currentForm && currentForm !== form) {
        delete currentForm.dataset.busy;
        setComposerLocked(currentForm, false, state.chatWorkspace);
      }
    }
    if (pendingAutoNext && state.me?.id === user) {
      const nextJob = pendingAutoNext;
      queueMicrotask(() => {
        if (state.me?.id !== user || activeTurnFor(threadId)) return;
        const targetWorkspace = state.chatWorkspace?.session?.id === threadId ? state.chatWorkspace : workspace;
        const targetForm = isViewingThread(threadId, user) ? chatForm() : null;
        const nextSelected = Array.isArray((draftFor(targetWorkspace) || {}).selected)
          ? (draftFor(targetWorkspace) || {}).selected
          : (nextJob.selected || []);
        const nextPayload = {
          message: nextJob.message,
          mode: nextJob.selection ? "auto" : (nextJob.mode || "auto"),
          material_ids: readableSourcesOnly(nextSelected).map((file) => file.id),
          ...(nextJob.selection ? { selection: nextJob.selection } : {})
        };
        toastMessage(
          nextJob.kind === "redirect" ? "已结合追加引导重新开始执行" : "正在按刚才追加的引导继续处理…",
          "info"
        );
        void runConversationTurn({
          workspace: targetWorkspace,
          form: targetForm,
          message: nextJob.message,
          payload: nextPayload,
          selected: nextSelected,
          draft: draftFor(targetWorkspace),
          allowAiImages,
          resume: false,
          rethrowToForm: false
        }).catch((err) => {
          toastMessage(cleanErrorMessage(err, "追加引导执行失败"), "error");
        });
      });
    }
  }
}

export async function sendMessage(form, rerender) {
  const workspace = state.chatWorkspace;
  if (clientExecutionRequired() && typeof window.officeDesktop?.runAgent !== "function") {
    throw new Error("请在桌面客户端中使用 AI 办理");
  }
  if (hasActiveUploadsFor(workspace)) {
    throw new Error("附件正在上传中，请稍候完成后再发送");
  }
  const message = form.elements.message.value.trim();
  if (!message) throw new Error("请输入工作要求");
  const draft = draftFor(workspace);
  const selected = draft.selected || [];
  const docSelected = readableSourcesOnly(selected);
  const payload = {
    message, mode: draft.selection ? "auto" : form.elements.mode.value, material_ids: docSelected.map((file) => file.id),
    ...(draft.selection ? { selection: draft.selection } : {})
  };
  if (clientExecutionRequired() && !selectedModelAvailable(workspace)) {
    throw new Error(workspace.models?.length ? "请先选择可用模型" : "暂无可用模型，请联系管理员开通。");
  }
  const allowAiImages = form.elements.allow_ai_images ? form.elements.allow_ai_images.checked === true : imageAutoPreference();
  await runConversationTurn({
    workspace,
    form,
    message,
    payload,
    selected,
    draft,
    allowAiImages,
    resume: false,
    rethrowToForm: true
  });
}

export async function resumeFailedTurn(triggerButton = null) {
  const workspace = state.chatWorkspace;
  const threadId = workspace?.session?.id;
  if (!workspace || !threadId) throw new Error("当前会话不可用");
  if (!clientExecutionRequired() || typeof window.officeDesktop?.runAgent !== "function") {
    throw new Error("请在桌面客户端中使用断点续跑");
  }
  if (activeTurnFor(threadId)) throw new Error("当前会话有一轮任务正在执行");
  const form = chatForm();
  if (form?.dataset.busy) throw new Error("请等待当前请求执行结束");

  const checkpoint = workspace.checkpoint || await loadCheckpointFor(threadId, state.me?.id);
  if (!checkpoint) {
    workspace.checkpoint = null;
    syncCheckpointCardDom(workspace);
    throw new Error("未找到可续跑的任务断点");
  }
  workspace.checkpoint = checkpoint;

  if (!selectedModelAvailable(workspace)) {
    const fallbackModel = checkpoint.model;
    const canFallback = fallbackModel === PIPELINE_MODEL
      ? workspace.pipelineAvailable === true
      : Boolean(fallbackModel && workspace.models?.includes(fallbackModel));
    if (canFallback) {
      workspace.selectedModel = fallbackModel;
    } else if (!checkpoint.hasCompletedResult) {
      throw new Error(workspace.models?.length ? "请先选择一个可用模型后再续跑" : "暂无可用模型，请联系管理员开通。");
    }
  }

  // 若服务端仍残留未正常结束的 running 孤儿轮次，续跑前自动安全解锁。
  const turns = workspace.session?.turns || [];
  const orphanRunning = [...turns].reverse().find((t) => t.status === "running" && !t.reply && !t.error && t.id);
  if (orphanRunning) {
    const failPath = orphanRunning.kind === "image"
      ? `/conversations/${encodeURIComponent(threadId)}/images/${encodeURIComponent(orphanRunning.id)}/fail`
      : `/conversations/${encodeURIComponent(threadId)}/turns/${encodeURIComponent(orphanRunning.id)}/fail`;
    try {
      await post(failPath, { message: "上次执行异常中断，已自动转入断点续跑" });
    } catch (err) {
      if (err?.status !== 409 && err?.status !== 404) throw err;
    }
  }

  const lastTurn = turns[turns.length - 1];
  const message = (checkpoint.message || lastTurn?.message || "").trim();
  if (!message) throw new Error("无法恢复原始工作要求，请手动重新发送");

  const draft = draftFor(workspace);
  const payload = {
    message,
    mode: checkpoint.selection ? "auto" : (checkpoint.mode || "auto"),
    material_ids: Array.isArray(checkpoint.materialIds) ? checkpoint.materialIds : [],
    ...(checkpoint.selection ? { selection: checkpoint.selection } : {})
  };
  if (triggerButton?.isConnected) {
    triggerButton.disabled = true;
    triggerButton.textContent = "正在恢复断点…";
  }
  const allowAiImages = form?.elements?.allow_ai_images
    ? form.elements.allow_ai_images.checked === true
    : (checkpoint.allowImageTool !== false);
  await runConversationTurn({
    workspace,
    form,
    message,
    payload,
    selected: draft.selected || [],
    draft,
    allowAiImages,
    resume: true,
    rethrowToForm: false
  });
}

export async function discardCheckpoint() {
  const workspace = state.chatWorkspace;
  const threadId = workspace?.session?.id;
  if (!workspace || !threadId) return;
  try {
    await window.officeDesktop?.clearCheckpoint?.(threadId);
  } catch {}
  workspace.checkpoint = null;
  document.querySelector("[data-checkpoint-card]")?.remove();
  toastMessage("已放弃任务断点", "info");
}

/* ================= 最近会话：拖拽删除、拖拽排序、置顶、重命名与右键菜单 ================= */
let sessionMenuEl = null;
let sessionDragInitialized = false;
let draggingSessionId = "";
let draggingSessionTitle = "";

function findConversationSummary(sessionId) {
  const list = state.chatWorkspace?.conversations || state.conversations || [];
  return list.find((c) => c.id === sessionId) || null;
}

function replaceConversationsList(nextList) {
  state.conversations = nextList;
  if (state.chatWorkspace) state.chatWorkspace.conversations = nextList;
  emitConversationsUpdated();
}

export function closeSessionContextMenu() {
  if (sessionMenuEl) {
    sessionMenuEl.remove();
    sessionMenuEl = null;
  }
}

export async function deleteConversation(sessionId) {
  if (!sessionId) return false;
  closeSessionContextMenu();
  if (activeTurnFor(sessionId)) {
    toastMessage("该会话正在执行中，请先停止生成后再删除", "error");
    return false;
  }
  const existing = findConversationSummary(sessionId);
  const title = existing?.title || "该会话";
  await api("/conversations/" + encodeURIComponent(sessionId), { method: "DELETE" });
  sessionCache.delete(sessionId);
  try { await window.officeDesktop?.clearCheckpoint?.(sessionId); } catch {}
  const prev = state.chatWorkspace?.conversations || state.conversations || [];
  const next = prev.filter((c) => c.id !== sessionId);
  replaceConversationsList(next);
  if (state.route[0] === "chat" && state.route[1] === sessionId) {
    if (state.chatWorkspace) {
      state.chatWorkspace.session = null;
      state.chatWorkspace.checkpoint = null;
    }
    go("chat");
  } else if (dialog.open && dialog.querySelector(".session-history-row")) {
    showConversationHistory();
  }
  toastMessage(`已删除会话「${title}」`, "success");
  return true;
}

export async function togglePinConversation(sessionId) {
  if (!sessionId) return;
  closeSessionContextMenu();
  const existing = findConversationSummary(sessionId);
  const nextPinned = !Boolean(existing?.pinned);
  await api("/conversations/" + encodeURIComponent(sessionId), {
    method: "PATCH",
    body: { pinned: nextPinned }
  });
  const refreshed = await api("/conversations");
  if (Array.isArray(refreshed)) replaceConversationsList(refreshed);
  if (dialog.open && dialog.querySelector(".session-history-row")) showConversationHistory();
  toastMessage(nextPinned ? "已置顶该会话" : "已取消置顶", "success");
}

export function openRenameConversationModal(sessionId) {
  if (!sessionId) return;
  closeSessionContextMenu();
  const existing = findConversationSummary(sessionId);
  const currentTitle = existing?.title || "";
  modal("重命名会话", `<form class="stack" data-form="session-rename" data-id="${e(sessionId)}">
    <label class="field">会话标题
      <input name="title" required maxlength="60" value="${e(currentTitle)}" placeholder="输入精炼的会话标题" autocomplete="off">
    </label>
    <div class="actions between">
      ${typeof window.officeDesktop?.generateTitle === "function"
        ? `<button type="button" class="text-button" data-act="session-ai-title" data-id="${e(sessionId)}">${icon("sparkles")}让 AI 自动命名</button>`
        : "<span></span>"}
      <div class="actions">
        <button type="button" class="button" data-act="close-dialog">取消</button>
        <button type="submit" class="button primary">保存名称</button>
      </div>
    </div>
  </form>`);
  const input = dialog.querySelector("form[data-form=session-rename] input[name=title]");
  if (input) {
    input.focus();
    input.select();
  }
}

export async function submitRenameConversation(form) {
  const sessionId = form.dataset.id;
  const title = String(new FormData(form).get("title") || "").trim();
  if (!sessionId || !title) throw new Error("请输入会话标题");
  const updated = await api("/conversations/" + encodeURIComponent(sessionId), {
    method: "PATCH",
    body: { title }
  });
  applyConversationTitleLocally(sessionId, updated.title || title);
  dialog.close();
  toastMessage("会话标题已更新", "success");
  return updated;
}

export async function retitleSingleConversation(sessionId) {
  if (!sessionId) return;
  closeSessionContextMenu();
  if (dialog.open && dialog.querySelector("form[data-form=session-rename]")) dialog.close();
  toastMessage("正在由 AI 生成精炼标题…", "info");
  const title = await generateConversationTitle(sessionId, { force: true });
  if (title) toastMessage(`已更新标题为「${title}」`, "success");
  else toastMessage("该会话暂无足够内容用于生成标题", "error");
}

export function openSessionContextMenu(sessionId, triggerOrPoint = null) {
  if (!sessionId) return;
  closeSessionContextMenu();
  const existing = findConversationSummary(sessionId);
  if (!existing) return;
  const pinned = Boolean(existing.pinned);
  const hasAiTitle = typeof window.officeDesktop?.generateTitle === "function";

  const menu = document.createElement("div");
  menu.className = "session-ctx-menu";
  menu.setAttribute("role", "menu");
  menu.innerHTML = `
    <button type="button" class="session-ctx-item" role="menuitem" data-act="session-pin" data-id="${e(sessionId)}">
      ${icon("pin")}<span>${pinned ? "取消置顶" : "置顶会话"}</span>
    </button>
    <button type="button" class="session-ctx-item" role="menuitem" data-act="session-rename" data-id="${e(sessionId)}">
      ${icon("edit")}<span>重命名</span>
    </button>
    ${hasAiTitle ? `<button type="button" class="session-ctx-item" role="menuitem" data-act="session-ai-title" data-id="${e(sessionId)}">
      ${icon("sparkles")}<span>AI 智能命名</span>
    </button>` : ""}
    <div class="session-ctx-divider" role="separator"></div>
    <button type="button" class="session-ctx-item is-danger" role="menuitem" data-act="session-delete" data-id="${e(sessionId)}">
      ${icon("trash")}<span>删除会话</span>
    </button>
  `;
  document.body.appendChild(menu);
  sessionMenuEl = menu;

  let x = 16;
  let y = 16;
  if (triggerOrPoint instanceof Element) {
    const rect = triggerOrPoint.getBoundingClientRect();
    x = rect.right + 4;
    y = rect.top;
  } else if (triggerOrPoint && typeof triggerOrPoint.clientX === "number") {
    x = triggerOrPoint.clientX + 2;
    y = triggerOrPoint.clientY + 2;
  }
  const menuRect = menu.getBoundingClientRect();
  if (x + menuRect.width > window.innerWidth - 10) {
    x = Math.max(10, window.innerWidth - menuRect.width - 10);
  }
  if (y + menuRect.height > window.innerHeight - 10) {
    y = Math.max(10, window.innerHeight - menuRect.height - 10);
  }
  menu.style.left = `${Math.round(x)}px`;
  menu.style.top = `${Math.round(y)}px`;
}

async function commitSessionReorder(draggedId, targetId, placeAfter) {
  if (!draggedId || !targetId || draggedId === targetId) return;
  const current = [...(state.chatWorkspace?.conversations || state.conversations || [])];
  const fromIdx = current.findIndex((c) => c.id === draggedId);
  const targetObj = current.find((c) => c.id === targetId);
  if (fromIdx < 0 || !targetObj) return;

  const [moved] = current.splice(fromIdx, 1);
  const nextPinned = Boolean(targetObj.pinned);
  const pinnedChanged = Boolean(moved.pinned) !== nextPinned;
  moved.pinned = nextPinned;

  const toIdx = current.findIndex((c) => c.id === targetId);
  if (toIdx < 0) return;
  current.splice(placeAfter ? toIdx + 1 : toIdx, 0, moved);
  replaceConversationsList(current);

  try {
    if (pinnedChanged) {
      await api("/conversations/" + encodeURIComponent(draggedId), {
        method: "PATCH",
        body: { pinned: nextPinned }
      });
    }
    const refreshed = await api("/conversations/reorder", {
      method: "PUT",
      body: { ids: current.map((c) => c.id) }
    });
    if (Array.isArray(refreshed)) replaceConversationsList(refreshed);
  } catch (error) {
    toastMessage(cleanErrorMessage(error, "保存会话排序失败"), "error");
    try {
      const fallback = await api("/conversations");
      if (Array.isArray(fallback)) replaceConversationsList(fallback);
    } catch {}
  }
}

function clearSessionDragHighlights() {
  document.querySelectorAll("[data-session-row].is-drop-before, [data-session-row].is-drop-after, [data-session-row].is-dragging")
    .forEach((el) => el.classList.remove("is-drop-before", "is-drop-after", "is-dragging"));
  const trash = document.querySelector("[data-session-trash-zone]");
  if (trash) {
    trash.classList.remove("is-over");
    const label = trash.querySelector("[data-session-trash-label]");
    if (label) label.textContent = "拖到此处删除会话";
  }
}

export function initSessionDragAndMenu() {
  if (sessionDragInitialized) return;
  sessionDragInitialized = true;

  document.addEventListener("contextmenu", (event) => {
    const row = event.target.closest?.("[data-session-row]");
    if (!row) return;
    event.preventDefault();
    const id = row.dataset.sessionId;
    if (id) openSessionContextMenu(id, event);
  });

  document.addEventListener("dblclick", (event) => {
    const link = event.target.closest?.("[data-session-row] .session-link");
    if (!link) return;
    const row = link.closest("[data-session-row]");
    const id = row?.dataset.sessionId;
    if (id) {
      event.preventDefault();
      openRenameConversationModal(id);
    }
  });

  document.addEventListener("pointerdown", (event) => {
    if (!sessionMenuEl) return;
    if (sessionMenuEl.contains(event.target) || event.target.closest?.("[data-act=session-menu]")) return;
    closeSessionContextMenu();
  }, true);

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeSessionContextMenu();
  });
  window.addEventListener("resize", closeSessionContextMenu);
  window.addEventListener("scroll", (event) => {
    if (sessionMenuEl && !sessionMenuEl.contains(event.target)) closeSessionContextMenu();
  }, true);

  document.addEventListener("dragstart", (event) => {
    const row = event.target.closest?.("[data-session-row]");
    if (!row) return;
    if (event.target.closest?.(".session-row-actions")) {
      event.preventDefault();
      return;
    }
    const id = row.dataset.sessionId || "";
    if (!id) return;
    closeSessionContextMenu();
    draggingSessionId = id;
    draggingSessionTitle = findConversationSummary(id)?.title || "";
    row.classList.add("is-dragging");
    document.body.classList.add("is-dragging-session");
    try {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("application/x-qo-session", id);
      event.dataTransfer.setData("text/plain", draggingSessionTitle || id);
    } catch {}
  });

  document.addEventListener("dragover", (event) => {
    if (!draggingSessionId) return;
    const trash = event.target.closest?.("[data-session-trash-zone]");
    if (trash) {
      event.preventDefault();
      try { event.dataTransfer.dropEffect = "move"; } catch {}
      document.querySelectorAll("[data-session-row].is-drop-before, [data-session-row].is-drop-after")
        .forEach((el) => el.classList.remove("is-drop-before", "is-drop-after"));
      trash.classList.add("is-over");
      const label = trash.querySelector("[data-session-trash-label]");
      if (label) label.textContent = draggingSessionTitle ? `松手删除「${draggingSessionTitle.slice(0, 12)}」` : "松手立即删除";
      return;
    }
    const row = event.target.closest?.("[data-session-row]");
    const activeTrash = document.querySelector("[data-session-trash-zone]");
    if (activeTrash) {
      activeTrash.classList.remove("is-over");
      const label = activeTrash.querySelector("[data-session-trash-label]");
      if (label) label.textContent = "拖到此处删除会话";
    }
    if (!row || row.dataset.sessionId === draggingSessionId) return;
    event.preventDefault();
    try { event.dataTransfer.dropEffect = "move"; } catch {}
    const rect = row.getBoundingClientRect();
    const after = event.clientY > rect.top + rect.height / 2;
    document.querySelectorAll("[data-session-row].is-drop-before, [data-session-row].is-drop-after")
      .forEach((el) => {
        if (el !== row) el.classList.remove("is-drop-before", "is-drop-after");
      });
    row.classList.toggle("is-drop-before", !after);
    row.classList.toggle("is-drop-after", after);
  });

  document.addEventListener("dragleave", (event) => {
    if (!draggingSessionId) return;
    const trash = event.target.closest?.("[data-session-trash-zone]");
    if (trash && !trash.contains(event.relatedTarget)) {
      trash.classList.remove("is-over");
      const label = trash.querySelector("[data-session-trash-label]");
      if (label) label.textContent = "拖到此处删除会话";
    }
  });

  document.addEventListener("drop", (event) => {
    if (!draggingSessionId) return;
    const id = draggingSessionId;
    const trash = event.target.closest?.("[data-session-trash-zone]");
    if (trash) {
      event.preventDefault();
      event.stopPropagation();
      draggingSessionId = "";
      draggingSessionTitle = "";
      document.body.classList.remove("is-dragging-session");
      clearSessionDragHighlights();
      void deleteConversation(id).catch((err) => toastMessage(cleanErrorMessage(err, "删除会话失败"), "error"));
      return;
    }
    const row = event.target.closest?.("[data-session-row]");
    if (row && row.dataset.sessionId && row.dataset.sessionId !== id) {
      event.preventDefault();
      event.stopPropagation();
      const targetId = row.dataset.sessionId;
      const rect = row.getBoundingClientRect();
      const after = event.clientY > rect.top + rect.height / 2;
      draggingSessionId = "";
      draggingSessionTitle = "";
      document.body.classList.remove("is-dragging-session");
      clearSessionDragHighlights();
      void commitSessionReorder(id, targetId, after);
    }
  });

  document.addEventListener("dragend", () => {
    if (!draggingSessionId && !document.body.classList.contains("is-dragging-session")) return;
    draggingSessionId = "";
    draggingSessionTitle = "";
    document.body.classList.remove("is-dragging-session");
    clearSessionDragHighlights();
  });
}
