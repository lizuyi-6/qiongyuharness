import { active, api, badge, button, deadline, e, fileMarkup, go, icon, modal, outputWarnings, post, readDraft, saveDraft,
  state, toastMessage, uploadMarkup, uploadSelected, uuid } from "./ui.js";
import { deckFiles, deckPanel, selectionChip } from "./slides.js";

const readable = new Set(["docx", "pptx", "pdf", "txt", "md"]);
const modelDraftKey = "chat/last-model";
const PIPELINE_MODEL = "qo:pipeline";
const imageAutoPreferenceKey = "qo/chat/allow-ai-images";
let usageRequestId = 0;

function imageAutoPreference() {
  try {
    const stored = localStorage.getItem(imageAutoPreferenceKey);
    return stored === null ? true : stored === "true";
  } catch {
    return true;
  }
}

export function saveImageAutoPreference(value) {
  try { localStorage.setItem(imageAutoPreferenceKey, value ? "true" : "false"); } catch {}
}

function clientExecutionRequired() {
  return state.engine?.client_execution === true;
}

function desktopUsageEnabled() {
  return clientExecutionRequired() && typeof window.officeDesktop?.usage === "function";
}

function formatUsageAmount(value, unit = "USD") {
  return new Intl.NumberFormat(undefined, { style: "currency", currency: unit, maximumFractionDigits: 2 }).format(value);
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
      rows.push({ label: `${labels[item.window]}额度`, used: item.used, limit: item.limit,
        remaining: item.remaining, resetAt: item.reset_at || "", resetFallback: "窗口未开始" });
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
        rows.push({ label, used, limit, remaining: Math.max(0, limit - used), resetAt,
          resetFallback: "上游未提供重置时间" });
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
    unlimited.textContent = "不限额";
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
    const progress = document.createElement("progress");
    progress.max = 100;
    progress.value = Math.max(0, Math.min(100, row.used / row.limit * 100));
    progress.setAttribute("aria-label", `${row.label}使用比例`);
    const amounts = document.createElement("p");
    amounts.className = "usage-amounts";
    amounts.textContent = `已用 ${formatUsageAmount(row.used, usage.unit || "USD")} / ${formatUsageAmount(row.limit, usage.unit || "USD")}`;
    const remaining = document.createElement("p");
    remaining.className = "usage-remaining";
    remaining.textContent = row.remaining <= 0 ?
      `已用尽，${row.resetAt ? `${formatUsageTime(row.resetAt)}恢复` : row.resetFallback}` :
      `剩余 ${formatUsageAmount(row.remaining, usage.unit || "USD")}`;
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
  if (workspace) workspace.usageState = result;
  if (content?.isConnected) renderUsageResult(content, result);
  return result;
}

export async function toggleDesktopUsage(button) {
  const thread = button.closest(".agent-thread");
  const panel = thread?.querySelector("[data-usage-panel]");
  if (!panel) return;
  const opening = panel.hidden;
  panel.hidden = !opening;
  button.setAttribute("aria-expanded", String(opening));
  if (opening) await refreshDesktopUsage(state.chatWorkspace, panel);
}

export async function loadDesktopModelState() {
  if (!clientExecutionRequired()) return {};
  if (typeof window.officeDesktop?.runAgent !== "function") return {models:[], image_models:[], selectedModel:"", modelNotice:""};

  const storedModel=readDraft(modelDraftKey)?.model || "";
  try {
    if (typeof window.officeDesktop.models !== "function") throw new Error("当前桌面客户端不支持模型列表");
    const result=await window.officeDesktop.models();
    const models=[...new Set((Array.isArray(result?.models)?result.models:[])
      .filter(model=>typeof model==="string" && model.trim()).map(model=>model.trim()))];
    const image_models=[...new Set((Array.isArray(result?.image_models)?result.image_models:[])
      .filter(model=>typeof model==="string" && /^gpt-image-/.test(model.trim())).map(model=>model.trim()))].sort();
    const pipelineAvailable=result?.pipeline?.available===true;
    const pipelineReason=typeof result?.pipeline?.reason==="string"?result.pipeline.reason:"";
    const storedModelAvailable=storedModel===PIPELINE_MODEL?pipelineAvailable:models.includes(storedModel);
    const invalidated=Boolean(storedModel && !storedModelAvailable);
    const selectedModel=invalidated?"":storedModel || (pipelineAvailable?PIPELINE_MODEL:"");
    if (invalidated) saveDraft(modelDraftKey,{model:""});
    const modelNotice=invalidated?"上次选择的模型已不可用，已清空，请重新选择。":
      result?.error || (models.length || pipelineAvailable?"":"暂无可用模型，请联系管理员开通。");
    return {models,image_models,selectedModel,modelNotice,pipelineAvailable,pipelineReason};
  } catch (error) {
    return {models:[],image_models:[],selectedModel:"",modelNotice:error.message || "模型列表获取失败，请重试。"};
  }
}

function selectedModelAvailable(workspace) {
  return workspace?.selectedModel===PIPELINE_MODEL?workspace.pipelineAvailable===true:
    Boolean(workspace?.selectedModel && workspace.models?.includes(workspace.selectedModel));
}

function draftFor(workspace) {
  return readDraft(workspace.draftKey) || {};
}

function writeDraft(workspace, form, fields = {}) {
  const previous = draftFor(workspace);
  const inputs = form ? {message:form.elements.message.value, mode:form.elements.mode.value} : {};
  saveDraft(workspace.draftKey, {...previous, ...inputs, ...fields});
}

export async function loadWorkspace(parts) {
  const [, scope, assignment] = parts;
  const [session, conversations, today, desktopModels] = await Promise.all([
    scope && !["task","project"].includes(scope) ? api("/conversations/" + encodeURIComponent(scope)) : null,
    api("/conversations"),
    api("/me/today"),
    loadDesktopModelState()
  ]);
  const bound = session?.assignment_id || (scope === "task" ? assignment : null);
  const group=session?.project_id || (scope==="project"?assignment:null);
  const [task,project]=await Promise.all([bound?api("/assignments/"+encodeURIComponent(bound)):null,
    group?api("/projects/"+encodeURIComponent(group)):null]);
  return {
    session, conversations, task, project, ...desktopModels,
    tasks:today.items.filter(task=>active(task.status)),
    draftKey:"chat/" + (session?.id || (task ? "task/" + task.id : project?"project/"+project.id:"new"))
  };
}

export function conversationNav(workspace) {
  if (!workspace) return "";
  return `<div class="workspace-nav">
    ${button("新建会话","route",{id:"chat",icon:"plus",cls:"new-session"})}
    ${workspace.conversations.length ? `<p class="nav-label">最近会话</p>
      ${button("最近会话","chat-history",{icon:"history",cls:"history-trigger"})}
      <div class="recent-sessions">${workspace.conversations.map((session)=>{
        const selected=state.route[0]==="chat" && state.route[1]===session.id;
        return `<button type="button" class="session-link ${selected?"active":""}" data-act="route" data-id="chat/${e(session.id)}"
          title="${e(session.title)}" ${selected?'aria-current="page"':""}>
          ${icon("message")}<span>${e(session.title)}</span></button>`;
      }).join("")}</div>` : ""}
  </div>`;
}

export function showConversationHistory() {
  const sessions=state.chatWorkspace?.conversations || [];
  modal("最近会话", `<div class="stack">${sessions.map((session)=>
    button(session.title,"route",{id:"chat/"+session.id,icon:"message"})).join("")}</div>`);
}

function turnView(turn) {
  const reply=turn.reply || turn.error || (turn.status === "running" ? "正在执行…" : "");
  const images=Array.isArray(turn.images)?turn.images:[];
  const pipeline=turn.pipeline;
  const pipelineLabel=pipeline?.verdict==="pass"?"验收通过":pipeline?.verdict==="fail"?"验收未通过":
    pipeline?.verdict==="skipped" && pipeline?.tier==="light"?"轻量任务·未验收":"";
  const reviewLabels={pass:"复核：通过",fail:"复核：未通过",not_needed:"复核：无需复核",unavailable:"复核：不可用"};
  return `<section class="conversation-turn">
    <div class="request-entry"><p>${turn.selection?`<small class="request-scope">第 ${turn.selection.slide} 页${turn.selection.shape_id?" · 文本框":""}</small>`:""}${e(turn.message)}</p></div>
    ${reply || images.length ? `<div class="response-entry">
      <p class="response-byline">${icon("mark")}<strong>办公助理</strong></p>
      ${turn.steps?.length ? `<details class="tool-trace"><summary>执行记录 · ${turn.steps.length} 步</summary>
        <ul>${turn.steps.map(step=>`<li>${icon("check")}${e(step)}</li>`).join("")}</ul></details>` : ""}
      <p class="${turn.error?"error":"response-copy"}">${e(reply)}</p>
      ${pipeline && pipelineLabel?`<div class="pipeline-result">
        <div class="pipeline-result-status"><span class="pipeline-badge ${pipeline.verdict==="fail"?"failed":pipeline.verdict==="skipped"?"skipped":"passed"}">${e(pipelineLabel)}</span>
          <span class="pipeline-review">${e(reviewLabels[pipeline.review] || "复核：未知")}</span></div>
        ${pipeline.verdict==="fail" && Array.isArray(pipeline.issues) && pipeline.issues.length?
          `<ul class="pipeline-issues">${pipeline.issues.map(issue=>`<li>${e(issue)}</li>`).join("")}</ul>`:""}
      </div>`:""}
      ${outputWarnings(turn.output_telemetry)}
      ${turn.file ? `<div class="session-file">${fileMarkup(turn.file,{edit:true})}</div>` : ""}
      ${images.length?`<div class="turn-images">${images.map(image=>{
        const file=image.file || {};
        const id=typeof file.id==="string"?file.id:"";
        const alt=typeof image.prompt==="string"?image.prompt:"生成的图片";
        const quality=image.quality==="low"?"草稿":image.quality==="high"?"精细":"标准";
        return `<figure class="turn-image">
          <img src="/api/v1/files/${encodeURIComponent(id)}/image" alt="${e(alt)}" loading="lazy">
          <figcaption><span>${e(image.width)}×${e(image.height)} · ${quality}质量 · ${e(image.model)}</span>
            <a href="/api/v1/files/${encodeURIComponent(id)}/download">下载</a></figcaption>
        </figure>`;
      }).join("")}</div>`:""}
    </div>` : ""}
  </section>`;
}

function historyView(workspace) {
  const session=workspace.session;
  return `${session?.has_earlier?button("加载更早的消息","chat-earlier"):""}${(session?.turns || []).map(turnView).join("")}`;
}

function welcome(workspace) {
  const task=workspace.task;
  const examples=task ? [
    ["先梳理交付要求","先梳理这项任务的交付要求，列出需要补充的资料。"],
    ["制作初稿",task.kind==="pptx"?"按任务要求生成一份可编辑 PPT，保留待确认的信息。":"按任务要求生成一份可编辑 Word 初稿，保留待确认的信息。"]
  ] : [
    ["整理会议纪要","帮我整理会议纪要，按议题、结论和后续行动展开。"],
    ["起草 Word 文档","帮我起草一份可编辑 Word 文档。"],
    ["制作汇报 PPT","帮我制作一份可编辑的汇报 PPT。"]
  ];
  return `<div class="workspace-empty">
    <h1>${task?"开始办理任务":"开始新的会话"}</h1>
    <p>${task ? e(task.title) : "交代要做的工作，或从右侧关联一项今日任务。"}</p>
    <div class="prompt-examples">${examples.map(([label,value])=>
        button(label,"chat-example",{id:value})).join("")}</div>
  </div>`;
}

function contextView(workspace) {
  const task=workspace.task;
  const project=workspace.project;
  return `<aside class="workspace-context" aria-label="办公上下文">
    ${deckPanel(workspace)}
    <h2>办公上下文</h2>
    ${project?`<section class="context-group"><p class="context-label">项目共享资料</p><h3>${e(project.name)}</h3>
      ${project.context.length?`<ul class="context-files">${project.context.slice(0,6).map(item=>`<li>${icon("clip")}${e(item.title)}</li>`).join("")}</ul>`:'<p class="hint">暂无共享资料</p>'}
      ${button("打开项目","route",{id:"projects/"+project.id,icon:"right"})}</section>`:""}
    ${task ? `<section class="context-group">
      <p class="context-label">关联任务</p>
      <h3>${e(task.title)}</h3>
      <div class="context-status">${badge(task.status)}<span>${e(deadline(task.due_at))}</span></div>
      ${task.description? `<p class="context-copy">${e(task.description)}</p>`:""}
      ${task.deliverable? `<h4>交付要求</h4><p class="context-copy">${e(task.deliverable)}</p>`:""}
      ${button(active(task.status)?"回到任务交付":"查看任务","open-task",{id:task.id,icon:"right"})}
    </section>` : project?"":`<section class="context-group">
      <p class="context-label">今日任务</p>
      ${workspace.tasks.length ? `<ul class="task-context-list">${workspace.tasks.map((item)=>`<li><button type="button"
        data-act="route" data-id="chat/task/${encodeURIComponent(item.id)}">
        <strong>${e(item.title)}</strong><span>${e(deadline(item.due_at))}</span></button></li>`).join("")}</ul>`
        : '<p class="hint">暂无待办理任务</p>'}
      ${button("查看全部任务","route",{id:"today",icon:"right"})}
    </section>`}
    ${task?.materials.length ? `<section class="context-group"><p class="context-label">任务资料</p><ul class="context-files">
      ${task.materials.map(file=>`<li>${icon("clip")}${e(file.name)}</li>`).join("")}</ul></section>` : ""}
  </aside>`;
}

function selectedFiles(selected) {
  return selected.map((file)=>`<span class="attached-file">${e(file.name)}
    <button type="button" data-act="chat-remove-source" data-id="${e(file.id)}" aria-label="移除 ${e(file.name)}">${icon("x")}</button>
  </span>`).join("");
}

function composer(workspace) {
  const draft=draftFor(workspace);
  const desktopRequired=clientExecutionRequired();
  const desktopAvailable=typeof window.officeDesktop?.runAgent === "function";
  const modelSelected=selectedModelAvailable(workspace);
  const blocked=desktopRequired && !desktopAvailable;
  const modelRequired=desktopRequired && desktopAvailable;
  const imageModels=workspace.image_models || [];
  const showImageControls=desktopRequired && desktopAvailable && imageModels.length>0;
  const showImageAuto=showImageControls && workspace.selectedModel!==PIPELINE_MODEL;
  if (workspace.task && !active(workspace.task.status)) {
    return `<div class="composer-wrap">${blocked?'<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>':""}
      <p class="notice">关联任务已完成或正在审核，可查看会话和文件。</p></div>`;
  }
  return `<div class="composer-wrap">
  <form class="chat-composer" data-form="conversation" aria-label="发送消息">
    ${blocked?'<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>':""}
    <div class="attached-files" data-chat-selection>${selectionChip(draft.selection)}</div>
    <div class="attached-files" data-chat-selected>${selectedFiles(draft.selected || [])}</div>
    <textarea required name="message" rows="2" maxlength="6000" aria-label="办公请求"
      placeholder="输入工作要求…" ${blocked?"disabled":""}>${e(draft.message || "")}</textarea>
    <div class="composer-actions">
      <div class="actions">
        ${uploadMarkup("添加附件")}
        ${button("引用资料","chat-sources")}
        <select name="mode" aria-label="输出方式">${Object.entries({auto:"自动选择",reply:"仅回复",docx:"Word",pptx:"PowerPoint"})
          .map(([mode,label])=>`<option value="${mode}" ${draft.mode===mode?"selected":""}>${label}</option>`).join("")}</select>
      </div>
      ${modelRequired?`<label class="chat-model-picker"><span>模型</span><select data-chat-model aria-label="执行模型"
        ${workspace.models?.length?"":"disabled"}>
        <option value="${PIPELINE_MODEL}" ${workspace.selectedModel===PIPELINE_MODEL?"selected":""} ${workspace.pipelineAvailable?"":"disabled"}>智能分工（理解→执行→验收）</option>
        <option value="" ${workspace.selectedModel?"":"selected"}>${workspace.models?.length?"请选择模型":"暂无可用模型"}</option>
        ${(workspace.models || []).map(model=>`<option value="${e(model)}" ${workspace.selectedModel===model?"selected":""}>${e(model)}</option>`).join("")}
      </select></label>`:""}
      ${showImageControls?`<div class="image-controls" aria-label="图片生成设置">
        <label><span>尺寸</span><select name="image_size" aria-label="生图尺寸">
          <option value="1024x1024">方形 1024×1024</option><option value="1536x1024">横版 1536×1024</option>
          <option value="1024x1536">竖版 1024×1536</option><option value="1536x864">16:9 1536×864</option>
        </select></label>
        <label><span>质量</span><select name="image_quality" aria-label="生图质量">
          <option value="low">草稿</option><option value="medium" selected>标准</option><option value="high">精细</option>
        </select></label>
        ${button("生成图片","chat-generate-image",{disabled:blocked})}
      </div>`:""}
      ${showImageAuto?`<label class="image-auto-option"><input type="checkbox" name="allow_ai_images" ${imageAutoPreference()?"checked":""}>
        <span>允许 AI 自动生图（每轮最多 2 张，按张计费）</span></label>`:""}
      <button class="button primary" type="submit" ${blocked || (modelRequired && !modelSelected)?"disabled":""}>${icon("arrow")}发送</button>
    </div>
    ${modelRequired && (workspace.modelNotice || workspace.pipelineReason)?`<p class="chat-model-status ${workspace.models?.length?"":"error"}" data-chat-model-status role="status">${e([workspace.modelNotice,workspace.pipelineReason].filter(Boolean).join(" "))}</p>`:""}
  </form>
  </div>`;
}

export function workspaceView(workspace) {
  const {session,task,project}=workspace;
  const hasTurns=Boolean(session?.turns.length);
  const title=session?.title || (task ? task.title : project?project.name:"新会话");
  const showUsage=desktopUsageEnabled();
  return `<div class="agent-workspace ${deckFiles(session).length?"has-deck":""}">
    <section class="agent-thread" aria-label="办公会话">
      <header class="thread-header">
        <span>${icon("message")}${e(title)}</span>
        <div class="thread-header-actions">
          ${state.engine?.engine==="demo" ? '<span class="thread-engine">本地演示</span>' :
            state.engine?.model ? `<span class="thread-engine">${e(state.engine.model)}</span>` : ""}
          ${showUsage?'<button class="button usage-toggle" type="button" data-act="usage-toggle" aria-expanded="false">模型额度</button>':""}
        </div>
      </header>
      ${showUsage?`<section class="usage-panel" data-usage-panel aria-label="模型额度" hidden>
        <div class="usage-panel-heading"><h2>模型额度</h2>
          <button class="button" type="button" data-act="usage-refresh">刷新</button></div>
        <div data-usage-content><p class="hint" role="status">打开面板后获取额度。</p></div>
      </section>`:""}
      <div class="thread-scroll">
        ${hasTurns?"":welcome(workspace)}
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

export function selectDesktopModel(select, workspace, statusSelector="[data-chat-model-status]") {
  const submit=select.form?.querySelector('button[type="submit"]');
  const allowed=select.value===PIPELINE_MODEL?workspace?.pipelineAvailable===true:workspace?.models?.includes(select.value);
  if (!allowed) {
    if (workspace) workspace.selectedModel="";
    if (submit) submit.disabled=true;
    return false;
  }
  workspace.selectedModel=select.value;
  if (submit) submit.disabled=false;
  workspace.modelNotice="";
  saveDraft(modelDraftKey,{model:select.value});
  const status=document.querySelector(statusSelector);
  if (status) {
    status.textContent=workspace.pipelineReason || "";
    status.hidden=!workspace.pipelineReason;
  }
  return true;
}

export function selectChatModel(select) {
  return selectDesktopModel(select,state.chatWorkspace);
}

export function useExample(value) {
  const form=chatForm();
  if (!form || form.dataset.busy) return;
  form.elements.message.value=value;
  form.elements.message.focus();
  saveWorkspaceInput(form);
}

function updateSources(workspace, change) {
  const form=chatForm();
  if (workspace !== state.chatWorkspace || !form) return;
  const selected=change(draftFor(workspace).selected || []);
  writeDraft(workspace, form, {selected});
  form.querySelector("[data-chat-selected]").innerHTML=selectedFiles(selected);
}

export function removeSource(id) {
  updateSources(state.chatWorkspace, sources=>sources.filter(file=>file.id!==id));
}

export async function attachFiles(files, input) {
  const workspace=state.chatWorkspace;
  if (chatForm()?.dataset.busy) throw new Error("请等待当前请求执行结束");
  const existing=draftFor(workspace).selected || [];
  if (existing.length+files.length>8) throw new Error("每轮最多添加 8 份参考资料");
  if (files.some(file=>!readable.has(file.name.split(".").pop().toLowerCase()))) {
    throw new Error("支持引用 Word、PowerPoint、PDF、TXT 和 Markdown");
  }
  const controls=chatForm()?.querySelectorAll("button[type=submit]");
  controls?.forEach(control=>{control.disabled=true;});
  try {
    for (const file of files) {
      const [added]=await uploadSelected([file], workspace.task?.id || "");
      updateSources(workspace, sources=>[...sources,added]);
    }
    if (files.length) toastMessage("附件已加入会话");
  } finally {
    controls?.forEach(control=>{control.disabled=false;});
    if (input) input.value="";
  }
}

export async function showSources() {
  const workspace=state.chatWorkspace;
  if (chatForm()?.dataset.busy) return;
  const personal=await api("/files");
  if (state.chatWorkspace!==workspace) return;
  const available=[...(workspace.task?.materials || []),
    ...personal.filter(file=>!file.assignment_id || file.assignment_id===workspace.task?.id)]
    .filter(file=>readable.has(file.extension));
  const selected=new Map((draftFor(workspace).selected || []).map(file=>[file.id,file]));
  for (const file of available) if (!selected.has(file.id)) selected.set(file.id,file);
  workspace.availableSources=selected;
  const chosen=new Set((draftFor(workspace).selected || []).map(file=>file.id));
  modal("引用资料", `<form class="stack" data-form="chat-sources">
    <div class="file-list">${[...selected.values()].map(file=>fileMarkup(file,{check:"ids",checked:chosen.has(file.id)})).join("") ||
      '<p class="hint">当前可引用的文件为空。可以从输入区添加附件。</p>'}</div>
    <div class="actions"><button class="button primary" type="submit">使用所选资料</button></div></form>`);
}

export function chooseSources(form) {
  const workspace=state.chatWorkspace;
  const ids=new FormData(form).getAll("ids");
  if (ids.length>8) throw new Error("每轮最多选择 8 份资料");
  updateSources(workspace, ()=>ids.map(id=>workspace.availableSources.get(id)).filter(Boolean));
}

export function scrollToLatest() {
  const container=document.querySelector(".thread-scroll");
  if (container) container.scrollTop=container.scrollHeight;
}

export async function loadEarlier() {
  const workspace=state.chatWorkspace;
  const session=workspace.session;
  if (!session?.has_earlier) return;
  const prior=await api("/conversations/"+encodeURIComponent(session.id)+"?offset="+session.turns.length);
  if (state.chatWorkspace!==workspace) return;
  const container=document.querySelector(".thread-scroll");
  const height=container.scrollHeight;
  const position=container.scrollTop;
  session.turns=[...prior.turns,...session.turns];
  session.has_earlier=prior.has_earlier;
  container.querySelector("[data-chat-turns]").innerHTML=historyView(workspace);
  container.scrollTop=position+container.scrollHeight-height;
}

async function readEvents(response, progress) {
  if (!response.ok) {
    const body=await response.json().catch(()=>({}));
    const error=new Error(typeof body.detail==="string"?body.detail:"发送失败，请检查连接后重试");
    error.retryAsNew=response.status<500 && response.status!==409;
    throw error;
  }
  const reader=response.body.getReader();
  const decoder=new TextDecoder();
  let buffered="", completed;
  function handle(line) {
    if (!line.trim()) return;
    const entry=JSON.parse(line);
    if (entry.type==="step") progress(entry.label);
    if (entry.type==="complete") completed=entry.turn;
    if (entry.type==="error") throw Object.assign(new Error(entry.message),{retryAsNew:true});
  }
  try {
    for (;;) {
      const {done,value}=await reader.read();
      buffered+=decoder.decode(value,{stream:!done});
      const lines=buffered.split("\n");
      buffered=lines.pop();
      for (const line of lines) handle(line);
      if (done) { handle(buffered); break; }
    }
  } finally { reader.releaseLock(); }
  if (!completed) throw new Error("连接中断。可刷新会话确认执行结果。");
  if (completed.status==="failed") throw Object.assign(new Error(completed.error || "本轮未完成"),{retryAsNew:true});
  return completed;
}

function startProgress(form, message) {
  document.querySelector(".workspace-empty")?.remove();
  const messages=document.querySelector("[data-chat-turns]");
  messages.insertAdjacentHTML("beforeend", `<section class="conversation-turn pending-turn">
    <div class="request-entry"><p>${e(message)}</p></div>
    <div class="response-entry">
      <p class="response-byline">${icon("mark")}<strong>办公助理</strong></p>
      <details class="tool-trace" data-progress-trace hidden><summary>执行记录 · 0 步</summary><ul></ul></details>
      <p class="running-note" role="status">准备执行…</p>
    </div></section>`);
  form.elements.message.value="";
  scrollToLatest();
  return messages.lastElementChild;
}

function appendProgressStep(progress, label) {
  if (!progress?.isConnected || typeof label!=="string" || !label) return;
  const trace=progress.querySelector("[data-progress-trace]");
  const list=trace?.querySelector("ul");
  if (!trace || !list) return;
  const item=document.createElement("li");
  item.textContent=label;
  list.append(item);
  trace.hidden=false;
  trace.querySelector("summary").textContent=`执行记录 · ${list.children.length} 步`;
}

export async function runDesktopAgent(job, onStep=()=>{}, afterRun=result=>result) {
  const desktop=window.officeDesktop;
  if (typeof desktop?.runAgent!=="function") throw new Error("请在桌面客户端中使用 AI 办理");
  let unsubscribe=()=>{};
  try {
    try {
      const stop=desktop.onAgentEvent?.(event=>{
        if (event?.type==="step") onStep(event.label);
      });
      if (typeof stop==="function") unsubscribe=stop;
    } catch {}
    const result=await desktop.runAgent(job);
    return await afterRun(result);
  } finally {
    try { unsubscribe(); } catch {}
  }
}

async function executeDesktopTurn(workspace, threadId, requestId, payload, progress, allowImageTool) {
  const desktop=window.officeDesktop;
  if (typeof desktop?.runAgent!=="function") throw new Error("请在桌面客户端中使用 AI 办理");
  const model=workspace.selectedModel;
  if (!selectedModelAvailable(workspace)) throw new Error(workspace.models?.length?"请先选择可用模型":"暂无可用模型，请联系管理员开通。");

  const base="/conversations/"+encodeURIComponent(threadId)+"/turns";
  let prepared;
  let phase="prepare";
  try {
    prepared=await post(base+"/prepare",{id:requestId,...payload});
    if (!prepared?.run_id) throw new Error("服务端未返回本轮运行编号");
    for (const step of prepared.steps || []) appendProgressStep(progress,step);
    const note=progress?.querySelector(".running-note");
    if (note) note.textContent="正在调用模型…";
    phase="agent";
    const job={system:prepared.system,prompt:prepared.prompt,context:prepared.context,model,
      ...(allowImageTool===true && workspace.image_models?.length && model!==PIPELINE_MODEL?{image_tool:true}:{})};
    return await runDesktopAgent(job,label=>{
      if (typeof label!=="string" || !label) return;
      appendProgressStep(progress,label);
      if (progress?.isConnected) progress.querySelector(".running-note").textContent=label;
    },async result=>{
      if (progress?.isConnected) progress.querySelector(".running-note").textContent="正在保存结果…";
      const images=Array.isArray(result.images)?result.images:[];
      let imageSaveFailures=0;
      for (const image of images) {
        try {
          await post(base+"/"+encodeURIComponent(prepared.run_id)+"/images",image);
        } catch {
          imageSaveFailures+=1;
        }
      }
      phase="complete";
      const completed=await post(base+"/"+encodeURIComponent(prepared.run_id)+"/complete",{
        text:result.text,usage:{...(result.usage || {}),max_tokens_clamped:result.max_tokens_clamped,
          finish_reasons:result.finish_reasons,truncated:result.truncated},model:result.model || model,
        ...(result.pipeline?{pipeline:result.pipeline}:{})
      });
      if (imageSaveFailures) completed.image_save_failures=imageSaveFailures;
      return completed;
    });
  } catch (error) {
    // turns/prepare never replays a request id (any existing turn with that id is a 409),
    // so reusing the pending id would block every retry of this message.
    error.retryAsNew=true;
    if (prepared?.run_id) {
      try {
        await post(base+"/"+encodeURIComponent(prepared.run_id)+"/fail",{message:error.message || "桌面执行失败"});
      } catch (failError) {
        if (phase==="complete" && failError.status===409 && (!error.status || error.status>=500)) {
          error.message="无法确认本轮是否已保存，请刷新会话查看最新结果。";
        } else if (failError.status!==409) error.message+=`（失败状态同步未成功：${failError.message}）`;
      }
    }
    throw error;
  }
}

function setImageProgress(node, value, error = false) {
  if (!node?.isConnected) return;
  const note=node.querySelector("[data-image-progress]");
  if (note) {
    note.textContent=value;
    note.classList.toggle("error",error);
  }
}

export async function generateImageFromComposer(button, rerender) {
  const workspace=state.chatWorkspace;
  const form=chatForm();
  if (!workspace || !form) throw new Error("当前会话不可用");
  if (form.dataset.busy) throw new Error("请等待当前请求执行结束");
  if (typeof window.officeDesktop?.generateImage!=="function") throw new Error("请在桌面客户端中生成图片");
  const prompt=form.elements.message.value.trim();
  if (!prompt) throw new Error("请先输入图片说明或提示词");
  const imageModels=workspace.image_models || [];
  const model=imageModels.includes("gpt-image-2.6-flare")?"gpt-image-2.6-flare":imageModels[0];
  if (!model) throw new Error("暂无可用的生图模型");
  const size=form.elements.image_size.value;
  const quality=form.elements.image_quality.value;
  const owner=state.me?.id;
  const threadId=workspace.session?.id || draftFor(workspace).newId || uuid();
  const turns=document.querySelector("[data-chat-turns]");
  const pending=document.createElement("section");
  pending.className="conversation-turn pending-turn";
  const request=document.createElement("div");
  request.className="request-entry";
  const requestText=document.createElement("p");
  requestText.textContent=prompt;
  request.append(requestText);
  const response=document.createElement("div");
  response.className="response-entry";
  const note=document.createElement("p");
  note.className="running-note";
  note.dataset.imageProgress="";
  note.setAttribute("role","status");
  note.textContent="正在生成图片…";
  response.append(note);
  pending.append(request,response);
  turns?.append(pending);
  scrollToLatest();
  form.dataset.busy="1";
  button.disabled=true;
  const originalLabel=button.textContent;
  button.textContent="正在生成图片…";
  let prepared;
  let generated;
  let saved=false;
  let phase="prepare";
  try {
    if (!workspace.session) {
      workspace.session=await post("/conversations",{id:threadId,assignment_id:workspace.task?.id || null,
        project_id:workspace.project?.id || null});
      if (state.chatWorkspace===workspace) {
        history.replaceState(null,"","#chat/"+encodeURIComponent(threadId));
        state.route=["chat",threadId];
      }
    }
    const base="/conversations/"+encodeURIComponent(threadId)+"/images";
    prepared=await post(base+"/prepare",{prompt,model,size,quality});
    if (!prepared?.run_id) throw new Error("服务端未返回生图运行编号");
    phase="generate";
    setImageProgress(pending,"正在生成图片…");
    generated=await window.officeDesktop.generateImage({prompt,model,size,quality});
    phase="complete";
    setImageProgress(pending,"正在保存图片…");
    const completed=await post(base+"/"+encodeURIComponent(prepared.run_id)+"/complete",{
      png_base64:generated.png_base64,width:generated.width,height:generated.height,
      model:generated.model,size:generated.size,quality:generated.quality,usage:generated.usage || {}
    });
    saved=true;
    if (state.chatWorkspace===workspace && state.me?.id===owner) await rerender(true);
    return completed;
  } catch (error) {
    if (prepared?.run_id && !saved) {
      const message=generated && phase==="complete" ? `图片已生成并计费，但保存失败：${error.message}` : error.message || "生图失败";
      try {
        await post("/conversations/"+encodeURIComponent(threadId)+"/images/"+encodeURIComponent(prepared.run_id)+"/fail",
          {message:message.slice(0,300)});
      } catch {}
      setImageProgress(pending,message,true);
    } else {
      setImageProgress(pending,error.message || "生图失败",true);
    }
    throw error;
  } finally {
    delete form.dataset.busy;
    if (button.isConnected) {
      button.disabled=false;
      button.textContent=originalLabel;
    }
  }
}

export async function sendMessage(form, rerender) {
  const workspace=state.chatWorkspace;
  const user=state.me.id;
  if (clientExecutionRequired() && typeof window.officeDesktop?.runAgent!=="function") {
    throw new Error("请在桌面客户端中使用 AI 办理");
  }
  const message=form.elements.message.value.trim();
  if (!message) throw new Error("请输入工作要求");
  const draft=draftFor(workspace);
  const selected=draft.selected || [];
  const payload={message,mode:draft.selection?"auto":form.elements.mode.value,material_ids:selected.map(file=>file.id),
    ...(draft.selection?{selection:draft.selection}:{})};
  const signature=JSON.stringify(payload);
  if (clientExecutionRequired() && !selectedModelAvailable(workspace)) {
    throw new Error(workspace.models?.length?"请先选择可用模型":"暂无可用模型，请联系管理员开通。");
  }
  const requestId=draft.pending?.signature===signature?draft.pending.id:uuid();
  const threadId=workspace.session?.id || draft.newId || uuid();
  writeDraft(workspace, form, {newId:threadId,pending:{id:requestId,signature}});
  const inView=()=>form.isConnected && state.me?.id===user;
  let progress;
  try {
    if (!workspace.session) {
      workspace.session=await post("/conversations",{id:threadId,assignment_id:workspace.task?.id || null,
        project_id:workspace.project?.id||null});
      if (inView()) {
        history.replaceState(null,"","#chat/"+encodeURIComponent(threadId));
        state.route=["chat",threadId];
      }
    }
    if (inView()) progress=startProgress(form,message);
    let completed;
    if (clientExecutionRequired()) {
      completed=await executeDesktopTurn(workspace,threadId,requestId,payload,progress,
        form.elements.allow_ai_images?.checked===true);
    } else {
      const response=await fetch("/api/v1/conversations/"+encodeURIComponent(threadId)+"/turns", {
        method:"POST",headers:{"Content-Type":"application/json"},credentials:"same-origin",
        body:JSON.stringify({id:requestId,...payload})
      });
      completed=await readEvents(response,label=>{
        if (progress?.isConnected) progress.querySelector(".running-note").textContent=label;
      });
    }
    if (state.me?.id!==user) return;
    writeDraft(workspace,null,{message:"",newId:null,pending:null,selection:null});
    const deck=completed.file?.extension==="pptx"?{fileId:completed.file.id,page:draft.selection?.slide||1}:draft.deck;
    saveDraft("chat/"+threadId,{selected,mode:payload.mode,deck});
    if (completed.file && workspace.task) {
      const key="submit/"+workspace.task.id;
      const submission=readDraft(key) || {};
      const files=new Set(submission.artifact_ids || []);
      if (completed.file.parent_id) files.delete(completed.file.parent_id);
      files.add(completed.file.id);
      saveDraft(key,{...submission,artifact_ids:[...files],key:uuid()});
    }
    if (inView()) {
      await rerender(true);
      if (completed.image_save_failures>0) {
        const latest=document.querySelector("[data-chat-turns] .conversation-turn:last-child .response-entry");
        if (latest) {
          const warning=document.createElement("p");
          warning.className="output-warnings";
          warning.textContent=`有 ${completed.image_save_failures} 张图片保存失败（已计费）`;
          latest.append(warning);
        }
      }
      chatForm()?.elements.message.focus({preventScroll:true});
    }
    void refreshDesktopUsage().catch(()=>{});
  } catch (error) {
    if (state.me?.id===user && inView()) {
      const current=progress?.querySelector(".running-note");
      if (current) { current.textContent=error.message; current.classList.add("error"); }
      form.elements.message.value=message;
      const key=workspace.session?.id ? "chat/"+threadId : workspace.draftKey;
      const pending=error.retryAsNew?null:{id:requestId,signature};
      saveDraft(key,{...draft,message,mode:payload.mode,newId:threadId,pending});
      workspace.draftKey=key;
    }
    throw error;
  }
}
