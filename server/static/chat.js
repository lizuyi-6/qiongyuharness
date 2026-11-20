import { active, api, badge, button, copyText, deadline, e, fileMarkup, icon, modal, outputWarnings, post, readDraft, saveDraft,
  state, toastMessage, uploadMarkup, uploadSelected, uuid } from "./ui.js";
import { deckFiles, deckPanel, refreshCanvasLive, selectDeckFile, selectionChip, setCanvasLiveState, studioArtifacts } from "./slides.js";
import { renderMarkdown } from "./markdown.js";

const readable = new Set(["docx", "pptx", "pdf", "txt", "md"]);
const modelDraftKey = "chat/last-model";
const PIPELINE_MODEL = "qo:pipeline";
const imageAutoPreferenceKey = "qo/chat/allow-ai-images";
let usageRequestId = 0;
let activeTurnCancel = null;

// AI 自行生图按张计费：缺省关闭，由员工本人勾选开启；按账号分开记忆，避免共用电脑时沿用他人的选择。
function imageAutoPreferenceStorageKey() {
  return state.me?.id ? `${imageAutoPreferenceKey}/${state.me.id}` : "";
}

export function imageAutoPreference() {
  try {
    const key = imageAutoPreferenceStorageKey();
    return Boolean(key) && localStorage.getItem(key) === "true";
  } catch {
    return false;
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

function desktopUsageEnabled() {
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

function formatUsageTime(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return "窗口未开始";
  return date.toLocaleString();
}

function usageRows(usage) {
  const rows = [];
  if (usage.mode === "quota_limited") {
    if (usage.quota && usage.quota.limit > 0) rows.push({
      label: "总额度", used: usage.quota.used, limit: usage.quota.limit,
      remaining: usage.quota.remaining, resetAt: "", resetFallback: "上游未提供重置时间"
    });
    const labels = { "5h": "5 小时", "1d": "日", "7d": "周" };
    for (const item of usage.rate_limits || []) {
      if (item.limit <= 0) continue;
      rows.push({
        label: `${labels[item.window]}额度`, used: item.used, limit: item.limit,
        remaining: item.remaining, resetAt: item.reset_at || "", resetFallback: "窗口未开始"
      });
    }
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
          resetFallback: "上游未提供重置时间"
        });
      }
    }
  }
  return rows;
}

function renderUsageResult(target, result) {
  target.replaceChildren();
  if (result?.state !== "available" || !result.usage) {
    const messages = {
      unauthorized: "模型凭证无效，请重新登录后重试。",
      unavailable: "暂时无法获取模型额度。",
      invalid: "额度数据格式无效。"
    };
    const notice = document.createElement("p");
    notice.className = "hint";
    notice.setAttribute("role", "status");
    notice.textContent = messages[result?.state] || messages.unavailable;
    target.append(notice);
    return;
  }

  const usage = result.usage;
  if (usage.balance !== undefined || usage.planName === "钱包余额") {
    const balance = document.createElement("p");
    balance.className = "usage-balance";
    balance.textContent = `钱包余额：${formatUsageAmount(usage.balance ?? usage.remaining, usage.unit || "USD")}`;
    target.append(balance);
    return;
  }

  const rows = usageRows(usage);
  if (!rows.length) {
    const unlimited = document.createElement("p");
    unlimited.className = "usage-unlimited";
    unlimited.textContent = "当前账户不限额";
    target.append(unlimited);
    return;
  }

  const list = document.createElement("div");
  list.className = "usage-grid";
  for (const row of rows) {
    const card = document.createElement("article");
    card.className = "usage-card";
    const title = document.createElement("h3");
    title.textContent = row.label;
    const ratio = Math.max(0, Math.min(100, (row.used / row.limit) * 100));
    const progress = document.createElement("progress");
    progress.max = 100;
    progress.value = ratio;
    if (ratio >= 92) progress.className = "danger";
    else if (ratio >= 75) progress.className = "warn";
    progress.setAttribute("aria-label", `${row.label}使用比例`);
    const amounts = document.createElement("p");
    amounts.className = "usage-amounts";
    amounts.textContent = `已用 ${formatUsageAmount(row.used, usage.unit || "USD")} / ${formatUsageAmount(row.limit, usage.unit || "USD")}`;
    const remaining = document.createElement("p");
    remaining.className = "usage-remaining";
    remaining.textContent = row.remaining <= 0
      ? `已用尽，${row.resetAt ? `${formatUsageTime(row.resetAt)}恢复` : row.resetFallback}`
      : `剩余 ${formatUsageAmount(row.remaining, usage.unit || "USD")}`;
    const reset = document.createElement("p");
    reset.className = "usage-reset";
    reset.textContent = row.resetAt ? `重置时间：${formatUsageTime(row.resetAt)}` : `重置时间：${row.resetFallback}`;
    card.append(title, progress, amounts, remaining, reset);
    list.append(card);
  }
  target.append(list);
}

export async function refreshDesktopUsage(workspace = state.chatWorkspace, panel = document.querySelector("[data-usage-panel]")) {
  if (!desktopUsageEnabled()) return null;
  const requestId = ++usageRequestId;
  const content = panel?.querySelector("[data-usage-content]");
  if (content) {
    content.replaceChildren();
    const loading = document.createElement("p");
    loading.className = "hint";
    loading.setAttribute("role", "status");
    loading.textContent = "正在获取模型额度…";
    content.append(loading);
  }
  let result;
  try { result = await window.officeDesktop.usage(); }
  catch { result = { state: "unavailable" }; }
  if (requestId !== usageRequestId) return result;
  state.usageState = result;
  if (workspace) workspace.usageState = result;
  if (content?.isConnected) renderUsageResult(content, result);
  return result;
}

export function syncDesktopUsagePanel() {
  if (!desktopUsageEnabled() || !state.usageOpen) return;
  const panel = document.querySelector("[data-usage-panel]");
  const content = panel?.querySelector("[data-usage-content]");
  if (!panel || !content) return;
  panel.hidden = false;
  if (state.usageState) renderUsageResult(content, state.usageState);
  else void refreshDesktopUsage(state.chatWorkspace, panel);
}

export async function toggleDesktopUsage(triggerButton) {
  const thread = triggerButton.closest(".agent-thread");
  const panel = thread?.querySelector("[data-usage-panel]");
  if (!panel) return;
  const opening = panel.hidden;
  state.usageOpen = opening;
  panel.hidden = !opening;
  triggerButton.setAttribute("aria-expanded", String(opening));
  if (opening) await refreshDesktopUsage(state.chatWorkspace, panel);
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

export async function loadDesktopModelState({ allowPipeline = true, draftKey = modelDraftKey } = {}) {
  if (!clientExecutionRequired()) return {};
  if (typeof window.officeDesktop?.runAgent !== "function") return { models: [], image_models: [], selectedModel: "", modelNotice: "" };

  const storedModel = readDraft(draftKey)?.model || "";
  try {
    if (typeof window.officeDesktop.models !== "function") throw new Error("当前桌面客户端不支持模型列表");
    const result = await window.officeDesktop.models();
    const models = [...new Set((Array.isArray(result?.models) ? result.models : [])
      .filter((model) => typeof model === "string" && model.trim()).map((model) => model.trim()))];
    const image_models = [...new Set((Array.isArray(result?.image_models) ? result.image_models : [])
      .filter((model) => typeof model === "string" && /^gpt-image-/.test(model.trim())).map((model) => model.trim()))].sort();
    const pipelineAvailable = allowPipeline && result?.pipeline?.available === true;
    const pipelineReason = allowPipeline && typeof result?.pipeline?.reason === "string" ? result.pipeline.reason : "";
    const storedModelAvailable = storedModel === PIPELINE_MODEL ? pipelineAvailable : models.includes(storedModel);
    const invalidated = Boolean(storedModel && !storedModelAvailable);
    const selectedModel = invalidated ? "" : storedModel || (pipelineAvailable ? PIPELINE_MODEL : "");
    if (invalidated) saveDraft(draftKey, { model: "" });
    const modelNotice = invalidated ? "上次选择的模型已不可用，已清空，请重新选择。" :
      result?.error || (models.length || pipelineAvailable ? "" : "暂无可用模型，请联系管理员开通。");
    return { models, image_models, selectedModel, modelNotice, pipelineAvailable, pipelineReason };
  } catch (error) {
    return { models: [], image_models: [], selectedModel: "", modelNotice: error.message || "模型列表获取失败，请重试。" };
  }
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

export function syncChatSubmit(form = chatForm(), workspace = state.chatWorkspace) {
  if (!form) return;
  const disabled = Boolean(form.dataset.busy) || !canSubmitChat(workspace);
  form.querySelectorAll('button[type="submit"]').forEach((control) => {
    control.disabled = disabled;
  });
}

function draftFor(workspace) {
  return readDraft(workspace.draftKey) || {};
}

function writeDraft(workspace, form, fields = {}) {
  const previous = draftFor(workspace);
  const inputs = form ? { message: form.elements.message.value, mode: form.elements.mode.value } : {};
  saveDraft(workspace.draftKey, { ...previous, ...inputs, ...fields });
}

export async function loadWorkspace(parts) {
  const [, scope, assignment] = parts;
  const [session, conversations, today, desktopModels] = await Promise.all([
    scope && !["task", "project"].includes(scope) ? api("/conversations/" + encodeURIComponent(scope)) : null,
    api("/conversations"),
    api("/me/today"),
    loadDesktopModelState()
  ]);
  const bound = session?.assignment_id || (scope === "task" ? assignment : null);
  const group = session?.project_id || (scope === "project" ? assignment : null);
  const [task, project] = await Promise.all([
    bound ? api("/assignments/" + encodeURIComponent(bound)) : null,
    group ? api("/projects/" + encodeURIComponent(group)) : null
  ]);
  state.conversations = conversations;
  state.today = today.items || [];
  if (today.today && state.me) state.me.today = today.today;
  return {
    session, conversations, task, project, ...desktopModels,
    tasks: (today.items || []).filter((item) => active(item.status)),
    draftKey: "chat/" + (session?.id || (task ? "task/" + task.id : project ? "project/" + project.id : "new"))
  };
}

export function conversationNav(workspace) {
  if (!workspace) return "";
  return `<div class="workspace-nav">
    <button type="button" class="button new-session" data-act="route" data-id="chat" title="新建会话">
      ${icon("plus")}<span class="new-session-label">新建会话</span>
    </button>
    ${workspace.conversations.length ? `<p class="nav-label">最近会话</p>
      <button type="button" class="button ghost history-trigger" data-act="chat-history" title="最近会话" aria-label="最近会话">${icon("history")}</button>
      <div class="recent-sessions">${workspace.conversations.map((session) => {
        const selected = state.route[0] === "chat" && state.route[1] === session.id;
        return `<button type="button" class="session-link ${selected ? "active" : ""}" data-act="route" data-id="chat/${e(session.id)}"
          title="${e(session.title)}" ${selected ? 'aria-current="page"' : ""}>
          ${icon("message")}<span>${e(session.title)}</span></button>`;
      }).join("")}</div>` : ""}
  </div>`;
}

export function showConversationHistory() {
  const sessions = state.chatWorkspace?.conversations || state.conversations || [];
  modal("最近会话", `<div class="stack-sm">${sessions.map((session) =>
    button(session.title, "route", { id: "chat/" + session.id, icon: "message" })).join("") || '<p class="hint">暂无历史会话</p>'}</div>`);
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

function turnView(turn, index, canRetry = false) {
  const reply = turn.reply || turn.error || (turn.status === "running" ? "正在执行…" : "");
  const images = Array.isArray(turn.images) ? turn.images : [];
  const pipeline = turn.pipeline;
  const pipelineLabel = pipeline?.verdict === "pass" ? "验收通过" : pipeline?.verdict === "fail" ? "验收未通过" :
    pipeline?.verdict === "skipped" && pipeline?.tier === "light" ? "轻量任务 · 未验收" : "";
  const reviewLabels = { pass: "复核：通过", fail: "复核：未通过", not_needed: "复核：无需复核", unavailable: "复核：不可用" };
  return `<section class="conversation-turn">
    <div class="request-entry"><p>${turn.selection ? `<small class="request-scope">第 ${turn.selection.slide} 页${turn.selection.shape_id ? " · 文本框" : " · 整页"}</small><br>` : ""}${e(turn.message)}</p></div>
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
      ${turn.error ? `<p class="error">${e(reply)}</p>` : `<div class="response-copy md-body">${renderMarkdown(reply)}</div>`}
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

function historyView(workspace) {
  const session = workspace.session;
  const canRetry = !(workspace.task && !active(workspace.task.status));
  return `${session?.has_earlier ? `<div class="actions">${button("加载更早的消息", "chat-earlier", { icon: "history", cls: "sm" })}</div>` : ""}${(session?.turns || []).map((turn, i) => turnView(turn, i, canRetry)).join("")}`;
}

function welcomeHero(workspace) {
  const { task, project } = workspace;
  const firstName = state.me?.name || "";
  const heading = task ? task.title : project ? project.name : `${firstName ? `${firstName}，` : ""}今天要处理什么？`;
  const tag = task ? "关联任务" : project ? "项目空间" : "";
  return `<div class="welcome-hero">
    <div class="welcome-emblem" aria-hidden="true">${icon("mark")}</div>
    ${tag ? `<span class="welcome-tag">${e(tag)}</span>` : ""}
    <h1>${e(heading)}</h1>
  </div>`;
}

function welcomeBelow(workspace) {
  const { task, project } = workspace;
  const examples = task ? [
    ["梳理交付要求", "先梳理这项任务的交付要求，列出需要补充的资料。"],
    ["生成初稿", task.kind === "pptx" ? "按任务要求生成一份可编辑 PPT，保留待确认的信息。" : "按任务要求生成一份可编辑 Word 初稿，保留待确认的信息。"]
  ] : [
    ["整理会议纪要", "帮我整理会议纪要，按议题、结论和后续行动展开。"],
    ["起草 Word 文档", "帮我起草一份可编辑 Word 文档。"],
    ["制作汇报 PPT", "帮我制作一份可编辑的汇报 PPT。"]
  ];

  const quickTasks = !task && !project && workspace.tasks?.length
    ? `<section class="welcome-tasks" aria-label="今日待办任务">
        <div class="between">
          <span class="context-label">今日待办</span>
          <button type="button" class="text-button" data-act="route" data-id="today">全部${icon("right")}</button>
        </div>
        <div class="welcome-task-grid">${workspace.tasks.slice(0, 4).map((item) => `
          <button type="button" class="welcome-task-item" data-act="route" data-id="chat/task/${encodeURIComponent(item.id)}">
            <strong>${e(item.title)}</strong>
            <span class="${item.overdue ? "late" : ""}">${e(deadline(item.due_at))}${item.overdue ? " · 逾期" : ""}</span>
          </button>`).join("")}
        </div>
      </section>`
    : "";

  return `<div class="welcome-below">
    <div class="prompt-examples">${examples.map(([label, value]) =>
      button(label, "chat-example", { id: value })).join("")}</div>
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
  return `<div class="context-anchor">
    <div class="context-anchor-main">
      <span class="context-anchor-tag">项目空间</span>
      <strong title="${e(project.name)}">${e(project.name)}</strong>
      <span class="hint">${project.context.length} 份共享资料</span>
    </div>
    <div class="context-anchor-actions">
      <button type="button" class="text-button" data-act="toggle-context" data-id="context">共享资料</button>
      <button type="button" class="text-button" data-act="route" data-id="projects/${e(project.id)}">打开项目${icon("right")}</button>
    </div>
  </div>`;
}

function contextBody(workspace) {
  const task = workspace.task;
  const project = workspace.project;
  return `<div class="studio-context-body">
    ${project ? `<section class="context-group"><p class="context-label">项目共享资料</p><h3>${e(project.name)}</h3>
      ${project.context.length ? `<ul class="context-files">${project.context.slice(0, 8).map((item) => `<li>${icon("clip")}<span>${e(item.title)}</span></li>`).join("")}</ul>` : '<p class="hint">暂无共享资料</p>'}
      ${button("打开项目空间", "route", { id: "projects/" + project.id, icon: "right", cls: "sm" })}</section>` : ""}
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

function selectedFiles(selected) {
  return selected.map((file) => `<span class="attached-file">${icon("clip")}<span>${e(file.name)}</span>
    <button type="button" data-act="chat-remove-source" data-id="${e(file.id)}" aria-label="移除 ${e(file.name)}">${icon("x")}</button>
  </span>`).join("");
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
  const showImageAuto = showImageControls && workspace.selectedModel !== PIPELINE_MODEL;
  const imageAutoOn = showImageAuto && imageAutoPreference();
  const drawerOpen = showImageControls && Boolean(state.imageDrawerOpen);
  const selectedCount = (draft.selected || []).length;
  const draftText = draft.message || "";
  if (workspace.task && !active(workspace.task.status)) {
    return `<div class="composer-wrap"><div class="chat-composer">${blocked ? '<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>' : ""}
      <p class="notice">关联任务已完成或正在审核，当前仅可查阅历史会话与文件。</p></div></div>`;
  }
  return `<div class="composer-wrap">
  <form class="chat-composer" data-form="conversation" aria-label="发送消息">
    <div class="slash-menu" data-slash-menu hidden role="listbox" aria-label="快捷指令"></div>
    ${blocked ? '<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>' : ""}
    <div class="attached-files" data-chat-selection>${selectionChip(draft.selection)}</div>
    <div class="attached-files" data-chat-selected>${selectedFiles(draft.selected || [])}</div>
    <div class="composer-grow" data-composer-grow data-replicated-value="${e(draftText + "\n")}">
      <textarea required name="message" rows="1" maxlength="6000" aria-label="办公请求"
        placeholder="描述你想完成的工作，输入 / 唤起指令…" ${blocked ? "disabled" : ""}>${e(draftText)}</textarea>
    </div>
    <div class="composer-actions">
      <div class="composer-tools">
        <label class="composer-tool upload-zone" data-upload-zone title="添加附件（也可直接拖拽文件到页面）" aria-label="添加附件">
          ${icon("clip")}<input type="file" data-upload accept=".docx,.pptx,.xlsx,.pdf,.txt,.md" multiple>
        </label>
        <button type="button" class="composer-tool ${selectedCount ? "has-value" : ""}" data-act="chat-sources" data-sources-btn title="引用已有资料" aria-label="引用资料">
          ${icon("folder")}<span class="tool-badge" data-sources-count ${selectedCount ? "" : "hidden"}>${selectedCount}</span>
        </button>
        <button type="button" class="composer-tool" data-act="slash-trigger" title="快捷指令（/）" aria-label="快捷指令">
          ${icon("slash")}
        </button>
        <span class="composer-sep" aria-hidden="true"></span>
        <select class="composer-select" name="mode" aria-label="输出方式" title="输出方式">${Object.entries({ auto: "自动", reply: "仅回复", docx: "Word", pptx: "PPT" })
          .map(([mode, label]) => `<option value="${mode}" ${draft.mode === mode ? "selected" : ""}>${label}</option>`).join("")}</select>
        ${showImageControls ? `<button type="button" class="composer-tool ${drawerOpen ? "active" : ""} ${imageAutoOn ? "dot-active" : ""}" data-act="toggle-image-drawer" data-image-toggle aria-expanded="${drawerOpen ? "true" : "false"}" title="图片生成设置" aria-label="图片生成设置">
          ${icon("image")}
        </button>` : ""}
      </div>
      <div class="composer-right">
        ${modelRequired ? `<select class="composer-select model-select" data-chat-model aria-label="执行模型" title="选择执行模型"
          ${workspace.models?.length ? "" : "disabled"}>
          <option value="${PIPELINE_MODEL}" ${workspace.selectedModel === PIPELINE_MODEL ? "selected" : ""} ${workspace.pipelineAvailable ? "" : "disabled"}>智能分工</option>
          <option value="" ${workspace.selectedModel ? "" : "selected"}>${workspace.models?.length ? "选择模型" : "暂无模型"}</option>
          ${(workspace.models || []).map((model) => `<option value="${e(model)}" ${workspace.selectedModel === model ? "selected" : ""}>${e(model)}</option>`).join("")}
        </select>` : ""}
        <span class="composer-enter-hint" aria-hidden="true">↵</span>
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
          <label class="image-auto-option" data-image-auto-wrap ${showImageAuto ? "" : "hidden"}><input type="checkbox" name="allow_ai_images" ${imageAutoPreference() ? "checked" : ""}>
            <span>允许 AI 自动生图（每轮最多 2 张，按张计费）</span></label>
        </div>
      </div>
    </div>` : ""}
    ${modelRequired && (workspace.modelNotice || workspace.pipelineReason) ? `<p class="chat-model-status ${workspace.models?.length ? "" : "error"}" data-chat-model-status role="status">${e([workspace.modelNotice, workspace.pipelineReason].filter(Boolean).join(" "))}</p>` : ""}
  </form>
  </div>`;
}

export function workspaceView(workspace) {
  const { session, task, project } = workspace;
  const hasTurns = Boolean(session?.turns.length);
  const title = session?.title || (task ? task.title : project ? project.name : "新会话");
  const showUsage = desktopUsageEnabled();
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
          <span>支持 Word、PowerPoint、PDF、TXT、Markdown</span>
        </div>
      </div>
      <header class="thread-header">
        <div class="thread-title ${hasTurns ? "" : "muted"}" data-thread-title-wrap>
          ${hasTurns ? `${icon("message")}<span class="thread-title-text" title="${e(title)}">${e(title)}</span>` : ""}
        </div>
        <div class="thread-header-actions">
          ${state.engine?.engine === "demo" ? '<span class="thread-engine">本地演示</span>' :
            state.engine?.model ? `<span class="thread-engine mono">${e(state.engine.model)}</span>` : ""}
          ${showUsage ? `<button class="button sm ghost usage-toggle" type="button" data-act="usage-toggle" aria-expanded="${state.usageOpen ? "true" : "false"}">额度</button>` : ""}
          <button class="button sm ghost studio-trigger ${open && activeTab === "deck" ? "active" : ""}" type="button" data-act="toggle-context" data-id="deck" aria-expanded="${open && activeTab === "deck" ? "true" : "false"}">${icon("pptx")}画布</button>
          <button class="button sm ghost studio-trigger ${open && activeTab === "context" ? "active" : ""}" type="button" data-act="toggle-context" data-id="context" aria-expanded="${open && activeTab === "context" ? "true" : "false"}">${icon("panelRight")}上下文</button>
        </div>
      </header>
      ${contextAnchorBar(workspace)}
      ${showUsage ? `<section class="usage-panel" data-usage-panel aria-label="模型额度" ${state.usageOpen ? "" : "hidden"}>
        <div class="usage-panel-heading"><h2>模型额度</h2>
          <button class="button sm" type="button" data-act="usage-refresh">${icon("refresh")}刷新</button></div>
        <div data-usage-content><p class="hint" role="status">打开面板后获取额度。</p></div>
      </section>` : ""}
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
  const wrap = form?.querySelector("[data-image-auto-wrap]");
  if (toggleBtn && checkbox) {
    toggleBtn.classList.toggle("dot-active", Boolean(checkbox.checked) && !wrap?.hidden);
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
    saveWorkspaceInput(form);
    pulseComposer();
    toastMessage(nameMap[id], "info");
    textarea.focus();
  } else if (id === "image") {
    toggleImageDrawer(true);
    pulseComposer();
    textarea.focus();
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
  const autoWrap = select.form?.querySelector("[data-image-auto-wrap]");
  if (autoWrap) {
    autoWrap.hidden = select.value === PIPELINE_MODEL;
    syncImageAutoIndicator();
  }
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

function updateSources(workspace, change) {
  const form = chatForm();
  if (workspace !== state.chatWorkspace || !form) return;
  const selected = change(draftFor(workspace).selected || []);
  writeDraft(workspace, form, { selected });
  form.querySelector("[data-chat-selected]").innerHTML = selectedFiles(selected);
  const countEl = form.querySelector("[data-sources-count]");
  const btnEl = form.querySelector("[data-sources-btn]");
  if (countEl) {
    countEl.textContent = String(selected.length);
    countEl.hidden = selected.length === 0;
  }
  if (btnEl) btnEl.classList.toggle("has-value", selected.length > 0);
}

export function removeSource(id) {
  updateSources(state.chatWorkspace, (sources) => sources.filter((file) => file.id !== id));
}

export async function attachFiles(files, input) {
  const workspace = state.chatWorkspace;
  if (chatForm()?.dataset.busy) throw new Error("请等待当前请求执行结束");
  const existing = draftFor(workspace).selected || [];
  if (existing.length + files.length > 8) throw new Error("每轮最多添加 8 份参考资料");
  if (files.some((file) => !readable.has(file.name.split(".").pop().toLowerCase()))) {
    throw new Error("支持引用 Word、PowerPoint、PDF、TXT 和 Markdown");
  }
  const controls = chatForm()?.querySelectorAll("button[type=submit]");
  controls?.forEach((control) => { control.disabled = true; });
  try {
    for (const file of files) {
      const [added] = await uploadSelected([file], workspace.task?.id || "");
      updateSources(workspace, (sources) => [...sources, added]);
    }
    if (files.length) toastMessage("附件已加入会话", "success");
  } finally {
    syncChatSubmit(chatForm(), workspace);
    if (input) input.value = "";
  }
}

export async function showSources() {
  const workspace = state.chatWorkspace;
  if (chatForm()?.dataset.busy) return;
  const personal = await api("/files");
  if (state.chatWorkspace !== workspace) return;
  const available = [...(workspace.task?.materials || []),
    ...personal.filter((file) => !file.assignment_id || file.assignment_id === workspace.task?.id)]
    .filter((file) => readable.has(file.extension));
  const selected = new Map((draftFor(workspace).selected || []).map((file) => [file.id, file]));
  for (const file of available) if (!selected.has(file.id)) selected.set(file.id, file);
  workspace.availableSources = selected;
  const chosen = new Set((draftFor(workspace).selected || []).map((file) => file.id));
  modal("引用资料", `<form class="stack" data-form="chat-sources">
    <div class="file-list">${[...selected.values()].map((file) => fileMarkup(file, { check: "ids", checked: chosen.has(file.id) })).join("") ||
      '<p class="hint">当前可引用的文件为空。可以从输入区直接添加附件。</p>'}</div>
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

function startProgress(form, message) {
  closeSlashMenu();
  leaveEmptyState();
  const messages = document.querySelector("[data-chat-turns]");
  const canCancel = !clientExecutionRequired() || typeof window.officeDesktop?.cancelAgent === "function";
  messages.insertAdjacentHTML("beforeend", `<section class="conversation-turn pending-turn">
    <div class="request-entry"><p>${e(message)}</p></div>
    <div class="response-entry">
      <div class="response-head">
        <p class="response-byline"><span class="response-mark is-working">${icon("mark")}</span><strong>穹驭</strong></p>
      </div>
      <ol class="live-steps" data-progress-trace hidden></ol>
      <div class="running-bar">
        <span class="running-pulse" aria-hidden="true"></span>
        <p class="running-note" role="status">正在思考…</p>
        ${canCancel ? `<button type="button" class="button stop-button" data-act="chat-cancel">${icon("stop")}停止</button>` : ""}
      </div>
    </div></section>`);
  form.elements.message.value = "";
  syncComposerGrow(form.elements.message);
  scrollToLatest();
  return messages.lastElementChild;
}

function setComposerLocked(form, locked, workspace = state.chatWorkspace) {
  if (!form?.isConnected) return;
  for (const field of form.querySelectorAll("textarea, select, input[type=file], button[data-act=chat-sources], button[data-act=chat-remove-source], button[data-act=slash-trigger]")) {
    field.disabled = locked;
  }
  if (!locked) {
    const modelSelect = form.querySelector("[data-chat-model]");
    if (modelSelect) modelSelect.disabled = !workspace?.models?.length;
    syncChatSubmit(form, workspace);
  }
}

export async function cancelActiveChat() {
  if (typeof activeTurnCancel !== "function") return;
  const cancel = activeTurnCancel;
  activeTurnCancel = null;
  const stopBtn = document.querySelector("[data-act=chat-cancel]");
  if (stopBtn) {
    stopBtn.disabled = true;
    stopBtn.textContent = "正在停止…";
  }
  await cancel();
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

export async function runDesktopAgent(job, onStep = () => {}, afterRun = (result) => result) {
  const desktop = window.officeDesktop;
  if (typeof desktop?.runAgent !== "function") throw new Error("请在桌面客户端中使用 AI 办理");
  let unsubscribe = () => {};
  try {
    try {
      const stop = desktop.onAgentEvent?.((event) => {
        if (event?.type === "step") onStep(event.label);
      });
      if (typeof stop === "function") unsubscribe = stop;
    } catch {}
    const result = await desktop.runAgent(job);
    return await afterRun(result);
  } finally {
    try { unsubscribe(); } catch {}
  }
}

async function executeDesktopTurn(workspace, threadId, requestId, payload, progress, allowImageTool, onLiveStep = () => {}) {
  const desktop = window.officeDesktop;
  if (typeof desktop?.runAgent !== "function") throw new Error("请在桌面客户端中使用 AI 办理");
  const model = workspace.selectedModel;
  if (!selectedModelAvailable(workspace)) throw new Error(workspace.models?.length ? "请先选择可用模型" : "暂无可用模型，请联系管理员开通。");

  const base = "/conversations/" + encodeURIComponent(threadId) + "/turns";
  let prepared;
  let phase = "prepare";
  let cancelled = false;
  const prepareAbort = new AbortController();
  activeTurnCancel = async () => {
    cancelled = true;
    if (phase === "prepare") prepareAbort.abort();
    else if (phase === "agent" && typeof desktop.cancelAgent === "function") {
      try { await desktop.cancelAgent(); } catch {}
    }
  };
  try {
    prepared = await api(base + "/prepare", { method: "POST", body: { id: requestId, ...payload }, signal: prepareAbort.signal });
    if (cancelled) throw new Error("已停止生成");
    if (!prepared?.run_id) throw new Error("服务端未返回本轮运行编号");
    for (const step of prepared.steps || []) {
      appendProgressStep(progress, step);
      onLiveStep(step);
    }
    const note = progress?.querySelector(".running-note");
    if (note) note.textContent = "正在调用模型…";
    onLiveStep("正在调用模型…");
    phase = "agent";
    const job = {
      system: prepared.system, prompt: prepared.prompt, context: prepared.context, model,
      ...(allowImageTool === true && workspace.image_models?.length && model !== PIPELINE_MODEL ? { image_tool: true } : {})
    };
    return await runDesktopAgent(job, (label) => {
      if (typeof label !== "string" || !label) return;
      appendProgressStep(progress, label);
      onLiveStep(label);
      if (progress?.isConnected) progress.querySelector(".running-note").textContent = label;
    }, async (result) => {
      const images = Array.isArray(result.images) ? result.images : [];
      let imageSaveFailures = 0;
      if (progress?.isConnected) {
        progress.querySelector(".running-note").textContent = "正在保存结果…";
        progress.querySelector("[data-act=chat-cancel]")?.remove();
      }
      onLiveStep("正在保存并同步画布…");
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
      phase = "complete";
      const completed = await post(base + "/" + encodeURIComponent(prepared.run_id) + "/complete", {
        text: result.text,
        usage: {
          ...(result.usage || {}), max_tokens_clamped: result.max_tokens_clamped,
          finish_reasons: result.finish_reasons, truncated: result.truncated
        },
        model: result.model || model,
        ...(result.pipeline ? { pipeline: result.pipeline } : {})
      });
      if (imageSaveFailures) completed.image_save_failures = imageSaveFailures;
      return completed;
    });
  } catch (error) {
    if ((cancelled || error?.name === "AbortError") && !String(error?.message || "").startsWith("已停止生成")) {
      error = new Error("已停止生成");
    }
    error.retryAsNew = true;
    if (prepared?.run_id) {
      try {
        await post(base + "/" + encodeURIComponent(prepared.run_id) + "/fail", { message: error.message || "桌面执行失败" });
      } catch (failError) {
        if (phase === "complete" && failError.status === 409 && (!error.status || error.status >= 500)) {
          error.message = "无法确认本轮是否已保存，请刷新会话查看最新结果。";
        } else if (failError.status !== 409) error.message += `（失败状态同步未成功：${failError.message}）`;
      }
    }
    throw error;
  } finally {
    activeTurnCancel = null;
  }
}

function setImageProgress(node, value, error = false) {
  if (!node?.isConnected) return;
  const note = node.querySelector("[data-image-progress]");
  if (note) {
    note.textContent = value;
    note.classList.toggle("error", error);
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
  const [freshSession, conversations] = await Promise.all([
    api("/conversations/" + encodeURIComponent(threadId)),
    api("/conversations").catch(() => null)
  ]);
  if (state.chatWorkspace !== workspace) return;
  workspace.session = freshSession;
  workspace.draftKey = "chat/" + threadId;
  if (Array.isArray(conversations)) {
    workspace.conversations = conversations;
    state.conversations = conversations;
    window.dispatchEvent(new CustomEvent("qo:conversations-updated"));
  }
  const turnsEl = document.querySelector("[data-chat-turns]");
  if (turnsEl) turnsEl.innerHTML = historyView(workspace);
  const titleWrap = document.querySelector("[data-thread-title-wrap]");
  if (titleWrap && freshSession?.title) {
    titleWrap.classList.remove("muted");
    titleWrap.innerHTML = `${icon("message")}<span class="thread-title-text" title="${e(freshSession.title)}">${e(freshSession.title)}</span>`;
  }
  const selChip = document.querySelector("[data-chat-selection]");
  if (selChip) selChip.innerHTML = "";
  scrollToLatest();

  const artifacts = studioArtifacts(freshSession);
  if (openCanvas || selectId) {
    state.contextOpen = true;
    state.studioTab = "deck";
    state._lastStudioKey = "chat/" + threadId;
    syncStudioDom();
    await refreshCanvasLive(workspace, { selectId: selectId || artifacts[0]?.id || "", preferredPage, flash: true });
  } else {
    setCanvasLiveState({ active: false });
    if (state.contextOpen && state.studioTab === "deck") {
      await refreshCanvasLive(workspace, { selectId: artifacts[0]?.id || "", preferredPage, flash: false });
    }
  }
}

export async function generateImageFromComposer(button, rerender) {
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
  leaveEmptyState();
  switchStudioTab("deck");
  setCanvasLiveState({ active: true, kind: "image", label: "正在构图并绘制图片…", prompt, steps: ["正在准备生图请求"] });
  const turns = document.querySelector("[data-chat-turns]");
  const pending = document.createElement("section");
  pending.className = "conversation-turn pending-turn";
  const request = document.createElement("div");
  request.className = "request-entry";
  const requestText = document.createElement("p");
  requestText.textContent = prompt;
  request.append(requestText);
  const response = document.createElement("div");
  response.className = "response-entry";
  const note = document.createElement("p");
  note.className = "running-note";
  note.dataset.imageProgress = "";
  note.setAttribute("role", "status");
  note.textContent = "正在生成图片…";
  response.append(note);
  pending.append(request, response);
  turns?.append(pending);
  scrollToLatest();
  form.dataset.busy = "1";
  button.disabled = true;
  const originalLabel = button.textContent;
  button.textContent = "正在生成图片…";
  let prepared;
  let generated;
  let saved = false;
  let phase = "prepare";
  try {
    if (!workspace.session) {
      workspace.session = await post("/conversations", {
        id: threadId, assignment_id: workspace.task?.id || null,
        project_id: workspace.project?.id || null
      });
      if (state.chatWorkspace === workspace) {
        history.replaceState(null, "", "#chat/" + encodeURIComponent(threadId));
        state.route = ["chat", threadId];
      }
    }
    const base = "/conversations/" + encodeURIComponent(threadId) + "/images";
    prepared = await post(base + "/prepare", { prompt, model, size, quality });
    if (!prepared?.run_id) throw new Error("服务端未返回生图运行编号");
    phase = "generate";
    setImageProgress(pending, "正在生成图片…");
    setCanvasLiveState({ active: true, kind: "image", label: "正在渲染高分辨率画面…", prompt, steps: ["已提交生图任务", "正在渲染画面"] });
    generated = await window.officeDesktop.generateImage({ prompt, model, size, quality });
    phase = "complete";
    setImageProgress(pending, "正在保存图片…");
    setCanvasLiveState({ active: true, kind: "image", label: "正在保存并同步到画布…", prompt, steps: ["已提交生图任务", "画面渲染完成", "正在写入画布"] });
    const completed = await post(base + "/" + encodeURIComponent(prepared.run_id) + "/complete", {
      png_base64: generated.png_base64, width: generated.width, height: generated.height,
      model: generated.model, size: generated.size, quality: generated.quality, usage: generated.usage || {}
    });
    saved = true;
    form.elements.message.value = "";
    syncComposerGrow(form.elements.message);
    if (state.chatWorkspace === workspace && state.me?.id === owner) {
      const latestImgId = completed?.images?.[0]?.file?.id || "";
      await refreshConversationInPlace(workspace, threadId, { selectId: latestImgId, openCanvas: true });
    }
    return completed;
  } catch (error) {
    setCanvasLiveState({ active: false });
    if (prepared?.run_id && !saved) {
      const message = generated && phase === "complete" ? `图片已生成并计费，但保存失败：${error.message}` : error.message || "生图失败";
      try {
        await post("/conversations/" + encodeURIComponent(threadId) + "/images/" + encodeURIComponent(prepared.run_id) + "/fail",
          { message: message.slice(0, 300) });
      } catch {}
      setImageProgress(pending, message, true);
    } else {
      setImageProgress(pending, error.message || "生图失败", true);
    }
    throw error;
  } finally {
    delete form.dataset.busy;
    if (button.isConnected) {
      button.disabled = false;
      button.textContent = originalLabel;
    }
  }
}

export async function sendMessage(form, rerender) {
  const workspace = state.chatWorkspace;
  const user = state.me.id;
  if (clientExecutionRequired() && typeof window.officeDesktop?.runAgent !== "function") {
    throw new Error("请在桌面客户端中使用 AI 办理");
  }
  const message = form.elements.message.value.trim();
  if (!message) throw new Error("请输入工作要求");
  const draft = draftFor(workspace);
  const selected = draft.selected || [];
  const payload = {
    message, mode: draft.selection ? "auto" : form.elements.mode.value, material_ids: selected.map((file) => file.id),
    ...(draft.selection ? { selection: draft.selection } : {})
  };
  const signature = JSON.stringify(payload);
  if (clientExecutionRequired() && !selectedModelAvailable(workspace)) {
    throw new Error(workspace.models?.length ? "请先选择可用模型" : "暂无可用模型，请联系管理员开通。");
  }
  const requestId = draft.pending?.signature === signature ? draft.pending.id : uuid();
  const threadId = workspace.session?.id || draft.newId || uuid();
  writeDraft(workspace, form, { newId: threadId, pending: { id: requestId, signature } });
  const inView = () => form.isConnected && state.me?.id === user;
  let progress;
  const liveKind = inferLiveCanvasKind(message, payload.mode, draft.selection, workspace);
  const liveSteps = [];
  const pushLiveStep = (label) => {
    if (!label) return;
    if (liveSteps[liveSteps.length - 1] !== label) liveSteps.push(label);
    if (liveKind || (state.contextOpen && state.studioTab === "deck")) {
      setCanvasLiveState({
        active: true,
        kind: liveKind || "pptx",
        label,
        prompt: message,
        steps: liveSteps,
        targetPage: draft.selection?.slide || null
      });
    }
  };
  setComposerLocked(form, true, workspace);
  try {
    if (!workspace.session) {
      workspace.session = await post("/conversations", {
        id: threadId, assignment_id: workspace.task?.id || null,
        project_id: workspace.project?.id || null
      });
      if (inView()) {
        history.replaceState(null, "", "#chat/" + encodeURIComponent(threadId));
        state.route = ["chat", threadId];
      }
    }
    if (inView()) {
      progress = startProgress(form, message);
      if (liveKind) {
        switchStudioTab("deck");
        pushLiveStep(draft.selection?.slide ? `正在准备修改第 ${draft.selection.slide} 页…` : "正在理解工作要求…");
      }
    }
    let completed;
    if (clientExecutionRequired()) {
      completed = await executeDesktopTurn(workspace, threadId, requestId, payload, progress,
        form.elements.allow_ai_images?.checked === true, pushLiveStep);
    } else {
      const controller = new AbortController();
      activeTurnCancel = async () => { controller.abort(); };
      try {
        const response = await fetch("/api/v1/conversations/" + encodeURIComponent(threadId) + "/turns", {
          method: "POST", headers: { "Content-Type": "application/json" }, credentials: "same-origin",
          signal: controller.signal,
          body: JSON.stringify({ id: requestId, ...payload })
        });
        completed = await readEvents(response, (label) => {
          appendProgressStep(progress, label);
          pushLiveStep(label);
          if (progress?.isConnected) progress.querySelector(".running-note").textContent = label;
        });
      } catch (error) {
        if (error?.name === "AbortError") throw Object.assign(new Error("已停止生成"), { retryAsNew: true });
        throw error;
      } finally {
        activeTurnCancel = null;
      }
    }
    if (state.me?.id !== user) return;
    writeDraft(workspace, null, { message: "", newId: null, pending: null, selection: null });
    const latestFileId = completed.file && ["pptx", "docx"].includes(completed.file.extension)
      ? completed.file.id
      : (completed.images?.[0]?.file?.id || "");
    const targetPage = completed.file?.extension === "pptx" ? (draft.selection?.slide || 1) : 1;
    const deck = latestFileId ? { fileId: latestFileId, page: targetPage } : draft.deck;
    saveDraft("chat/" + threadId, { selected, mode: payload.mode, deck });
    if (completed.file && workspace.task) {
      const key = "submit/" + workspace.task.id;
      const submission = readDraft(key) || {};
      const files = new Set(submission.artifact_ids || []);
      if (completed.file.parent_id) files.delete(completed.file.parent_id);
      files.add(completed.file.id);
      saveDraft(key, { ...submission, artifact_ids: [...files], key: uuid() });
    }
    if (inView()) {
      await refreshConversationInPlace(workspace, threadId, {
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
    }
    void refreshDesktopUsage().catch(() => {});
  } catch (error) {
    setCanvasLiveState({ active: false });
    if (state.me?.id === user && inView()) {
      progress?.querySelector("[data-act=chat-cancel]")?.remove();
      progress?.querySelector(".running-pulse")?.remove();
      progress?.querySelector(".response-mark")?.classList.remove("is-working");
      const current = progress?.querySelector(".running-note");
      if (current) { current.textContent = error.message; current.classList.add("error"); }
      form.elements.message.value = message;
      syncComposerGrow(form.elements.message);
      const key = workspace.session?.id ? "chat/" + threadId : workspace.draftKey;
      const pending = error.retryAsNew ? null : { id: requestId, signature };
      saveDraft(key, { ...draft, message, mode: payload.mode, newId: threadId, pending });
      workspace.draftKey = key;
    }
    throw error;
  } finally {
    activeTurnCancel = null;
    setComposerLocked(form, false, workspace);
  }
}
