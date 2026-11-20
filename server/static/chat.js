import { active, api, badge, button, copyText, deadline, e, fileMarkup, icon, modal, outputWarnings, post, readDraft, saveDraft,
  state, toastMessage, uploadMarkup, uploadSelected, uuid } from "./ui.js";
import { deckFiles, deckPanel, selectionChip } from "./slides.js";
import { renderMarkdown } from "./markdown.js";

const readable = new Set(["docx", "pptx", "pdf", "txt", "md"]);
const modelDraftKey = "chat/last-model";
const PIPELINE_MODEL = "qo:pipeline";
let usageRequestId = 0;
let activeTurnCancel = null;

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

export function toggleWorkspaceContext(triggerButton) {
  const workspaceEl = document.querySelector(".agent-workspace");
  if (!workspaceEl) return;
  state.contextOpen = !state.contextOpen;
  workspaceEl.classList.toggle("context-hidden", !state.contextOpen);
  if (triggerButton) {
    triggerButton.setAttribute("aria-expanded", String(state.contextOpen));
    triggerButton.title = state.contextOpen ? "收起右侧上下文" : "展开右侧上下文";
  }
}

export async function loadDesktopModelState({ allowPipeline = true, draftKey = modelDraftKey } = {}) {
  if (!clientExecutionRequired()) return {};
  if (typeof window.officeDesktop?.runAgent !== "function") return { models: [], selectedModel: "", modelNotice: "" };

  const storedModel = readDraft(draftKey)?.model || "";
  try {
    if (typeof window.officeDesktop.models !== "function") throw new Error("当前桌面客户端不支持模型列表");
    const result = await window.officeDesktop.models();
    const models = [...new Set((Array.isArray(result?.models) ? result.models : [])
      .filter((model) => typeof model === "string" && model.trim()).map((model) => model.trim()))];
    const pipelineAvailable = allowPipeline && result?.pipeline?.available === true;
    const pipelineReason = allowPipeline && typeof result?.pipeline?.reason === "string" ? result.pipeline.reason : "";
    const storedModelAvailable = storedModel === PIPELINE_MODEL ? pipelineAvailable : models.includes(storedModel);
    const invalidated = Boolean(storedModel && !storedModelAvailable);
    const selectedModel = invalidated ? "" : storedModel || (pipelineAvailable ? PIPELINE_MODEL : "");
    if (invalidated) saveDraft(draftKey, { model: "" });
    const modelNotice = invalidated ? "上次选择的模型已不可用，已清空，请重新选择。" :
      result?.error || (models.length || pipelineAvailable ? "" : "暂无可用模型，请联系管理员开通。");
    return { models, selectedModel, modelNotice, pipelineAvailable, pipelineReason };
  } catch (error) {
    return { models: [], selectedModel: "", modelNotice: error.message || "模型列表获取失败，请重试。" };
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

function turnView(turn, index) {
  const reply = turn.reply || turn.error || (turn.status === "running" ? "正在执行…" : "");
  const pipeline = turn.pipeline;
  const pipelineLabel = pipeline?.verdict === "pass" ? "验收通过" : pipeline?.verdict === "fail" ? "验收未通过" :
    pipeline?.verdict === "skipped" && pipeline?.tier === "light" ? "轻量任务 · 未验收" : "";
  const reviewLabels = { pass: "复核：通过", fail: "复核：未通过", not_needed: "复核：无需复核", unavailable: "复核：不可用" };
  return `<section class="conversation-turn">
    <div class="request-entry"><p>${turn.selection ? `<small class="request-scope">第 ${turn.selection.slide} 页${turn.selection.shape_id ? " · 文本框" : " · 整页"}</small><br>` : ""}${e(turn.message)}</p></div>
    ${reply ? `<div class="response-entry">
      <div class="response-head">
        <p class="response-byline">${icon("mark")}<strong>办公助理</strong></p>
        ${turn.reply && !turn.error ? `<button type="button" class="text-button copy-button" data-act="copy-reply" data-id="${index}" aria-label="复制回复">${icon("copy")}<span>复制</span></button>` : ""}
      </div>
      ${turn.steps?.length ? `<details class="tool-trace"><summary>执行记录 · ${turn.steps.length} 步</summary>
        <ul>${turn.steps.map((step) => `<li>${icon("check")}${e(step)}</li>`).join("")}</ul></details>` : ""}
      ${turn.error ? `<p class="error">${e(reply)}</p>` : `<div class="response-copy md-body">${renderMarkdown(reply)}</div>`}
      ${pipeline && pipelineLabel ? `<div class="pipeline-result">
        <div class="pipeline-result-status"><span class="pipeline-badge ${pipeline.verdict === "fail" ? "failed" : pipeline.verdict === "skipped" ? "skipped" : "passed"}">${e(pipelineLabel)}</span>
          <span class="pipeline-review">${e(reviewLabels[pipeline.review] || "复核：未知")}</span></div>
        ${pipeline.verdict === "fail" && Array.isArray(pipeline.issues) && pipeline.issues.length ?
          `<ul class="pipeline-issues">${pipeline.issues.map((issue) => `<li>${e(issue)}</li>`).join("")}</ul>` : ""}
      </div>` : ""}
      ${outputWarnings(turn.output_telemetry)}
      ${turn.file ? `<div class="session-file">${fileMarkup(turn.file, { edit: true })}</div>` : ""}
    </div>` : ""}
  </section>`;
}

function historyView(workspace) {
  const session = workspace.session;
  return `${session?.has_earlier ? `<div class="actions">${button("加载更早的消息", "chat-earlier", { icon: "history", cls: "sm" })}</div>` : ""}${(session?.turns || []).map((turn, i) => turnView(turn, i)).join("")}`;
}

function welcome(workspace) {
  const task = workspace.task;
  const project = workspace.project;
  const examples = task ? [
    ["先梳理交付要求", "先梳理这项任务的交付要求，列出需要补充的资料。"],
    ["制作初稿", task.kind === "pptx" ? "按任务要求生成一份可编辑 PPT，保留待确认的信息。" : "按任务要求生成一份可编辑 Word 初稿，保留待确认的信息。"]
  ] : [
    ["整理会议纪要", "帮我整理会议纪要，按议题、结论和后续行动展开。"],
    ["起草 Word 文档", "帮我起草一份可编辑 Word 文档。"],
    ["制作汇报 PPT", "帮我制作一份可编辑的汇报 PPT。"]
  ];

  const heading = task ? "开始办理任务" : project ? project.name : "开始新的会话";
  const subtitle = task
    ? task.title
    : project
      ? "当前会话已关联项目共享资料，可直接交代工作要求。"
      : "交代要做的工作，或从右侧关联一项今日任务。";

  return `<div class="workspace-empty">
    <h1>${e(heading)}</h1>
    <p class="welcome-sub">${e(subtitle)}</p>
    <div class="prompt-examples">${examples.map(([label, value]) =>
      button(label, "chat-example", { id: value })).join("")}</div>
  </div>`;
}

function contextView(workspace) {
  const task = workspace.task;
  const project = workspace.project;
  return `<aside class="workspace-context" aria-label="办公上下文">
    ${deckPanel(workspace)}
    <div class="context-head">
      <h2>办公上下文</h2>
      <button type="button" class="icon-button" data-act="toggle-context" aria-label="收起办公上下文" title="收起侧栏">${icon("x")}</button>
    </div>
    ${project ? `<section class="context-group"><p class="context-label">项目共享资料</p><h3>${e(project.name)}</h3>
      ${project.context.length ? `<ul class="context-files">${project.context.slice(0, 6).map((item) => `<li>${icon("clip")}<span>${e(item.title)}</span></li>`).join("")}</ul>` : '<p class="hint">暂无共享资料</p>'}
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
  if (workspace.task && !active(workspace.task.status)) {
    return `<div class="composer-wrap"><div class="chat-composer">${blocked ? '<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>' : ""}
      <p class="notice">关联任务已完成或正在审核，当前仅可查阅历史会话与文件。</p></div></div>`;
  }
  return `<div class="composer-wrap">
  <form class="chat-composer" data-form="conversation" aria-label="发送消息">
    ${blocked ? '<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>' : ""}
    <div class="attached-files" data-chat-selection>${selectionChip(draft.selection)}</div>
    <div class="attached-files" data-chat-selected>${selectedFiles(draft.selected || [])}</div>
    <textarea required name="message" rows="2" maxlength="6000" aria-label="办公请求"
      placeholder="输入工作要求…（Ctrl + Enter 发送）" ${blocked ? "disabled" : ""}>${e(draft.message || "")}</textarea>
    <div class="composer-actions">
      <div class="actions">
        ${uploadMarkup("添加附件")}
        ${button("引用资料", "chat-sources")}
        <select name="mode" aria-label="输出方式">${Object.entries({ auto: "自动选择", reply: "仅回复", docx: "Word", pptx: "PowerPoint" })
          .map(([mode, label]) => `<option value="${mode}" ${draft.mode === mode ? "selected" : ""}>${label}</option>`).join("")}</select>
      </div>
      <div class="composer-right">
        ${modelRequired ? `<label class="chat-model-picker"><span>模型</span><select data-chat-model aria-label="执行模型"
          ${workspace.models?.length ? "" : "disabled"}>
          <option value="${PIPELINE_MODEL}" ${workspace.selectedModel === PIPELINE_MODEL ? "selected" : ""} ${workspace.pipelineAvailable ? "" : "disabled"}>智能分工（理解→执行→验收）</option>
          <option value="" ${workspace.selectedModel ? "" : "selected"}>${workspace.models?.length ? "请选择模型" : "暂无可用模型"}</option>
          ${(workspace.models || []).map((model) => `<option value="${e(model)}" ${workspace.selectedModel === model ? "selected" : ""}>${e(model)}</option>`).join("")}
        </select></label>` : ""}
        <button class="button primary" type="submit" ${blocked || (modelRequired && !modelSelected) ? "disabled" : ""}>${icon("arrow")}发送</button>
      </div>
    </div>
    ${modelRequired && (workspace.modelNotice || workspace.pipelineReason) ? `<p class="chat-model-status ${workspace.models?.length ? "" : "error"}" data-chat-model-status role="status">${e([workspace.modelNotice, workspace.pipelineReason].filter(Boolean).join(" "))}</p>` : ""}
  </form>
  </div>`;
}

export function workspaceView(workspace) {
  const { session, task, project } = workspace;
  const hasTurns = Boolean(session?.turns.length);
  const title = session?.title || (task ? task.title : project ? project.name : "新会话");
  const showUsage = desktopUsageEnabled();
  const hasDeck = deckFiles(session).length > 0;
  if (hasDeck) state.contextOpen = true;

  return `<div class="agent-workspace ${hasDeck ? "has-deck" : ""} ${state.contextOpen ? "" : "context-hidden"}">
    <section class="agent-thread" aria-label="办公会话">
      <header class="thread-header">
        <div class="thread-title">
          ${icon("message")}
          <span class="thread-title-text" title="${e(title)}">${e(title)}</span>
        </div>
        <div class="thread-header-actions">
          ${state.engine?.engine === "demo" ? '<span class="thread-engine">本地演示</span>' :
            state.engine?.model ? `<span class="thread-engine mono">${e(state.engine.model)}</span>` : ""}
          ${showUsage ? `<button class="button sm usage-toggle" type="button" data-act="usage-toggle" aria-expanded="${state.usageOpen ? "true" : "false"}">模型额度</button>` : ""}
          <button class="icon-button" type="button" data-act="toggle-context" aria-expanded="${state.contextOpen ? "true" : "false"}" aria-label="切换右侧上下文" title="${state.contextOpen ? "收起办公上下文" : "展开办公上下文"}">${icon("panelRight")}</button>
        </div>
      </header>
      ${showUsage ? `<section class="usage-panel" data-usage-panel aria-label="模型额度" ${state.usageOpen ? "" : "hidden"}>
        <div class="usage-panel-heading"><h2>模型额度</h2>
          <button class="button sm" type="button" data-act="usage-refresh">${icon("refresh")}刷新</button></div>
        <div data-usage-content><p class="hint" role="status">打开面板后获取额度。</p></div>
      </section>` : ""}
      <div class="thread-scroll">
        ${hasTurns ? "" : welcome(workspace)}
        <div class="thread-content" data-chat-turns>${historyView(workspace)}</div>
      </div>
      ${composer(workspace)}
    </section>
    ${contextView(workspace)}
  </div>`;
}

function chatForm() {
  return document.querySelector("form[data-form=conversation]");
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
  return true;
}

export function selectChatModel(select) {
  return selectDesktopModel(select, state.chatWorkspace);
}

export function useExample(value) {
  const form = chatForm();
  if (!form || form.dataset.busy) return;
  form.elements.message.value = value;
  form.elements.message.focus();
  saveWorkspaceInput(form);
}

function updateSources(workspace, change) {
  const form = chatForm();
  if (workspace !== state.chatWorkspace || !form) return;
  const selected = change(draftFor(workspace).selected || []);
  writeDraft(workspace, form, { selected });
  form.querySelector("[data-chat-selected]").innerHTML = selectedFiles(selected);
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

function startProgress(form, message) {
  document.querySelector(".workspace-empty")?.remove();
  const messages = document.querySelector("[data-chat-turns]");
  const canCancel = !clientExecutionRequired() || typeof window.officeDesktop?.cancelAgent === "function";
  messages.insertAdjacentHTML("beforeend", `<section class="conversation-turn pending-turn">
    <div class="request-entry"><p>${e(message)}</p></div>
    <div class="response-entry">
      <p class="response-byline">${icon("mark")}<strong>办公助理</strong></p>
      <details class="tool-trace" data-progress-trace hidden><summary>执行记录 · 0 步</summary><ul></ul></details>
      <div class="running-bar">
        <p class="running-note" role="status">准备执行…</p>
        ${canCancel ? `<button type="button" class="button stop-button" data-act="chat-cancel">${icon("stop")}停止生成</button>` : ""}
      </div>
    </div></section>`);
  form.elements.message.value = "";
  scrollToLatest();
  return messages.lastElementChild;
}

function setComposerLocked(form, locked, workspace = state.chatWorkspace) {
  if (!form?.isConnected) return;
  for (const field of form.querySelectorAll("textarea, select, input[type=file], button[data-act=chat-sources], button[data-act=chat-remove-source]")) {
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
  const trace = progress.querySelector("[data-progress-trace]");
  const list = trace?.querySelector("ul");
  if (!trace || !list) return;
  const item = document.createElement("li");
  item.innerHTML = `${icon("check")}<span>${e(label)}</span>`;
  list.append(item);
  trace.hidden = false;
  trace.querySelector("summary").textContent = `执行记录 · ${list.children.length} 步`;
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

async function executeDesktopTurn(workspace, threadId, requestId, payload, progress) {
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
    for (const step of prepared.steps || []) appendProgressStep(progress, step);
    const note = progress?.querySelector(".running-note");
    if (note) note.textContent = "正在调用模型…";
    phase = "agent";
    return await runDesktopAgent({ system: prepared.system, prompt: prepared.prompt, context: prepared.context, model }, (label) => {
      if (typeof label !== "string" || !label) return;
      appendProgressStep(progress, label);
      if (progress?.isConnected) progress.querySelector(".running-note").textContent = label;
    }, (result) => {
      if (cancelled) throw new Error("已停止生成");
      if (progress?.isConnected) {
        progress.querySelector(".running-note").textContent = "正在保存结果…";
        progress.querySelector("[data-act=chat-cancel]")?.remove();
      }
      phase = "complete";
      return post(base + "/" + encodeURIComponent(prepared.run_id) + "/complete", {
        text: result.text,
        usage: {
          ...(result.usage || {}), max_tokens_clamped: result.max_tokens_clamped,
          finish_reasons: result.finish_reasons, truncated: result.truncated
        },
        model: result.model || model,
        ...(result.pipeline ? { pipeline: result.pipeline } : {})
      });
    });
  } catch (error) {
    if (cancelled || error?.name === "AbortError") {
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
    if (inView()) progress = startProgress(form, message);
    let completed;
    if (clientExecutionRequired()) {
      completed = await executeDesktopTurn(workspace, threadId, requestId, payload, progress);
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
    const deck = completed.file?.extension === "pptx" ? { fileId: completed.file.id, page: draft.selection?.slide || 1 } : draft.deck;
    saveDraft("chat/" + threadId, { selected, mode: payload.mode, deck });
    if (completed.file?.extension === "pptx") state.contextOpen = true;
    if (completed.file && workspace.task) {
      const key = "submit/" + workspace.task.id;
      const submission = readDraft(key) || {};
      const files = new Set(submission.artifact_ids || []);
      if (completed.file.parent_id) files.delete(completed.file.parent_id);
      files.add(completed.file.id);
      saveDraft(key, { ...submission, artifact_ids: [...files], key: uuid() });
    }
    if (inView()) {
      await rerender(true);
      chatForm()?.elements.message.focus({ preventScroll: true });
    }
    void refreshDesktopUsage().catch(() => {});
  } catch (error) {
    if (state.me?.id === user && inView()) {
      progress?.querySelector("[data-act=chat-cancel]")?.remove();
      progress?.querySelector(".running-pulse")?.remove();
      const current = progress?.querySelector(".running-note");
      if (current) { current.textContent = error.message; current.classList.add("error"); }
      form.elements.message.value = message;
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
