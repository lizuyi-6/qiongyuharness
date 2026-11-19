import { active, api, badge, button, deadline, e, fileMarkup, go, icon, modal, post, readDraft, saveDraft,
  state, toastMessage, uploadMarkup, uploadSelected, uuid } from "./ui.js";
import { deckFiles, deckPanel, selectionChip } from "./slides.js";

const readable = new Set(["docx", "pptx", "pdf", "txt", "md"]);
const modelDraftKey = "chat/last-model";

function clientExecutionRequired() {
  return state.engine?.client_execution === true;
}

async function loadDesktopModelState() {
  if (!clientExecutionRequired()) return {};
  if (typeof window.officeDesktop?.runAgent !== "function") return {models:[], selectedModel:"", modelNotice:""};

  const storedModel=readDraft(modelDraftKey)?.model || "";
  try {
    if (typeof window.officeDesktop.models !== "function") throw new Error("当前桌面客户端不支持模型列表");
    const result=await window.officeDesktop.models();
    const models=[...new Set((Array.isArray(result?.models)?result.models:[])
      .filter(model=>typeof model==="string" && model.trim()).map(model=>model.trim()))];
    const invalidated=Boolean(storedModel && !models.includes(storedModel));
    const selectedModel=invalidated?"":storedModel;
    if (invalidated) saveDraft(modelDraftKey,{model:""});
    const modelNotice=invalidated?"上次选择的模型已不可用，已清空，请重新选择。":
      result?.error || (models.length?"":"暂无可用模型，请联系管理员开通。");
    return {models,selectedModel,modelNotice};
  } catch (error) {
    return {models:[],selectedModel:"",modelNotice:error.message || "模型列表获取失败，请重试。"};
  }
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
  return `<section class="conversation-turn">
    <div class="request-entry"><p>${turn.selection?`<small class="request-scope">第 ${turn.selection.slide} 页${turn.selection.shape_id?" · 文本框":""}</small>`:""}${e(turn.message)}</p></div>
    ${reply ? `<div class="response-entry">
      <p class="response-byline">${icon("mark")}<strong>办公助理</strong></p>
      ${turn.steps?.length ? `<details class="tool-trace"><summary>执行记录 · ${turn.steps.length} 步</summary>
        <ul>${turn.steps.map(step=>`<li>${icon("check")}${e(step)}</li>`).join("")}</ul></details>` : ""}
      <p class="${turn.error?"error":"response-copy"}">${e(reply)}</p>
      ${turn.file ? `<div class="session-file">${fileMarkup(turn.file,{edit:true})}</div>` : ""}
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
  const modelSelected=Boolean(workspace.selectedModel && workspace.models?.includes(workspace.selectedModel));
  const blocked=desktopRequired && !desktopAvailable;
  const modelRequired=desktopRequired && desktopAvailable;
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
        <option value="">${workspace.models?.length?"请选择模型":"暂无可用模型"}</option>
        ${(workspace.models || []).map(model=>`<option value="${e(model)}" ${workspace.selectedModel===model?"selected":""}>${e(model)}</option>`).join("")}
      </select></label>`:""}
      <button class="button primary" type="submit" ${blocked || (modelRequired && !modelSelected)?"disabled":""}>${icon("arrow")}发送</button>
    </div>
    ${modelRequired && workspace.modelNotice?`<p class="chat-model-status ${workspace.models?.length?"":"error"}" data-chat-model-status role="status">${e(workspace.modelNotice)}</p>`:""}
  </form>
  </div>`;
}

export function workspaceView(workspace) {
  const {session,task,project}=workspace;
  const hasTurns=Boolean(session?.turns.length);
  const title=session?.title || (task ? task.title : project?project.name:"新会话");
  return `<div class="agent-workspace ${deckFiles(session).length?"has-deck":""}">
    <section class="agent-thread" aria-label="办公会话">
      <header class="thread-header">
        <span>${icon("message")}${e(title)}</span>
        ${state.engine?.engine==="demo" ? '<span class="thread-engine">本地演示</span>' :
          state.engine?.model ? `<span class="thread-engine">${e(state.engine.model)}</span>` : ""}
      </header>
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

export function selectChatModel(select) {
  const workspace=state.chatWorkspace;
  if (!workspace?.models?.includes(select.value)) return;
  workspace.selectedModel=select.value;
  workspace.modelNotice="";
  saveDraft(modelDraftKey,{model:select.value});
  const status=document.querySelector("[data-chat-model-status]");
  if (status) { status.hidden=true; status.textContent=""; }
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

async function executeDesktopTurn(workspace, threadId, requestId, payload, progress) {
  const desktop=window.officeDesktop;
  if (typeof desktop?.runAgent!=="function") throw new Error("请在桌面客户端中使用 AI 办理");
  const model=workspace.selectedModel;
  if (!workspace.models?.includes(model)) throw new Error(workspace.models?.length?"请先选择可用模型":"暂无可用模型，请联系管理员开通。");

  const base="/conversations/"+encodeURIComponent(threadId)+"/turns";
  let prepared;
  let phase="prepare";
  let unsubscribe=()=>{};
  try {
    prepared=await post(base+"/prepare",{id:requestId,...payload});
    if (!prepared?.run_id) throw new Error("服务端未返回本轮运行编号");
    for (const step of prepared.steps || []) appendProgressStep(progress,step);
    const note=progress?.querySelector(".running-note");
    if (note) note.textContent="正在调用模型…";
    try {
      const stop=desktop.onAgentEvent?.(event=>{
        if (event?.type!=="step") return;
        appendProgressStep(progress,event.label);
        if (progress?.isConnected && event.label) progress.querySelector(".running-note").textContent=event.label;
      });
      if (typeof stop==="function") unsubscribe=stop;
    } catch {}
    phase="agent";
    const result=await desktop.runAgent({system:prepared.system,prompt:prepared.prompt,context:prepared.context,model});
    if (progress?.isConnected) progress.querySelector(".running-note").textContent="正在保存结果…";
    phase="complete";
    return await post(base+"/"+encodeURIComponent(prepared.run_id)+"/complete",{
      text:result.text,usage:result.usage,model:result.model || model
    });
  } catch (error) {
    if (prepared?.run_id) {
      try {
        await post(base+"/"+encodeURIComponent(prepared.run_id)+"/fail",{message:error.message || "桌面执行失败"});
        error.retryAsNew=true;
      } catch (failError) {
        if (error.status && error.status<500 && error.status!==409) error.retryAsNew=true;
        if (phase==="complete" && failError.status===409 && (!error.status || error.status>=500)) {
          error.message="无法确认本轮是否已保存，请刷新会话查看最新结果。";
        } else if (failError.status!==409) error.message+=`（失败状态同步未成功：${failError.message}）`;
      }
    }
    throw error;
  } finally {
    try { unsubscribe(); } catch {}
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
  if (clientExecutionRequired() && !workspace.models?.includes(workspace.selectedModel)) {
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
      completed=await executeDesktopTurn(workspace,threadId,requestId,payload,progress);
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
      chatForm()?.elements.message.focus({preventScroll:true});
    }
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
