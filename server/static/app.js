import { active, api, applyTheme, button, clearDrafts, copyText, dialog, e, errorIn, fileMarkup, getTheme, go, icon, initCustomSelects, isManager, isReview, modal,
  post, readDraft, roles, root, saveDraft, state, toastMessage, toggleTheme, uploadSelected, uuid } from "./ui.js";
import { agentSetup, dashboard, demoUsers, detectDesktopVersion, filesView, historyView, loginView, pageHead, planSection,
  renderFilesList, settingsView, taskDetail, views } from "./views.js";
import { ensureShell, openCommandPalette, setMobileNav, syncConnectionIndicator, syncShell,
  syncSidebarConversations, syncThemeToggle, syncTodayBadge, syncUpdateUi, toggleSidebar, triggerTopbarSweep, triggerViewEnter, unmountShell } from "./shell.js";
import { assignmentValues, auditPage, changeAssignment, clearAdmin, saveTaskForm, showPublish, showTaskForm,
  taskAdmin } from "./admin.js";
import { abortOrphanTurn, applySlashCommand, attachFiles, autoTitleRecentConversations, cancelActiveChat, chooseSources, closeComposerPickers,
  closeSlashMenu, copyReply, discardCheckpoint, generateImageFromComposer, handleSlashKeydown, loadDesktopModelState, loadEarlier,
  loadWorkspace, openDeckInStudio, openImageLightbox, pickComposerOption, prefetchDesktopModels, refreshDesktopUsage,
  removeSource, resetChatState, resumeFailedTurn, retryTurn, runDesktopAgent, saveImageAutoPreference,
  saveWorkspaceInput, scrollToLatest, selectChatModel, selectDesktopModel, sendMessage, showConversationHistory,
  showSources, switchStudioTab, syncActiveChatView, syncChatSubmit, syncComposerGrow, syncDesktopUsagePanel,
  syncImageAutoIndicator, syncShellCanvasState, toggleComposerPicker, toggleDesktopUsage, toggleImageDrawer,
  toggleWorkspaceContext, triggerSlashMenu, updateSlashMenu, useExample, workspaceView } from "./chat.js";
import { projectDialog, projectsView, submitProject, updateInvite } from "./project.js";
import { clearSelection, mountDeck } from "./slides.js";

let historyOffset = 0;
let agentTask = null;
let agentFiles = [];
let navigation = 0;
let currentMode = "";
let timerBusy = false;
let agentDesktopModels = {};
let updateListenerAttached = false;
let lastUpdateCheckAt = 0;
const historySize = 20;

function applyDesktopUpdateState(next) {
  if (!next || typeof next !== "object") return;
  const prevStatus = state.update?.status;
  state.update = next;
  if (["available", "downloading", "ready"].includes(next.status) && next.status !== prevStatus) {
    state.updateBannerDismissed = false;
  }
  syncUpdateUi();
  if (next.appliedNotice?.version) {
    const modeLabel = next.appliedNotice.mode === "full" ? "完整更新" : "增量更新";
    toastMessage(`已通过${modeLabel}升级至 v${next.appliedNotice.version}`, "success");
    void window.officeDesktop?.clearUpdateNotice?.().catch(() => {});
  }
}

async function checkUpdateFallback({ silent = false } = {}) {
  lastUpdateCheckAt = Date.now();
  const current = detectDesktopVersion();
  if (!silent) {
    state.update = { ...(state.update || {}), status: "checking" };
    syncUpdateUi();
  }
  try {
    const query = current ? `?current=${encodeURIComponent(current)}` : "";
    const data = await api("/updates/latest" + query);
    state.update = {
      status: data?.has_update ? "available" : "up-to-date",
      hasUpdate: Boolean(data?.has_update),
      latestVersion: data?.latest_version || "",
      mode: data?.update_mode || "none",
      notes: data?.notes || "",
      publishedAt: data?.published_at || "",
      delta: data?.delta || null,
      full: data?.full || null,
      size: data?.full?.size || data?.delta?.size || 0,
      fullSize: data?.full?.size || 0
    };
    if (data?.has_update) state.updateBannerDismissed = false;
    syncUpdateUi();
    if (!silent) {
      toastMessage(data?.has_update ? `发现新版本 v${data.latest_version}` : "当前已是最新版本", data?.has_update ? "info" : "success");
    }
  } catch (error) {
    if (!silent) {
      state.update = { ...(state.update || {}), status: "error", error: error.message || "检查更新失败" };
      syncUpdateUi();
      toastMessage(error.message || "检查更新失败", "error");
    }
  }
}

async function initDesktopUpdater() {
  const desktop = window.officeDesktop;
  if (typeof desktop?.getUpdateState === "function") {
    if (!updateListenerAttached && typeof desktop.onUpdateEvent === "function") {
      updateListenerAttached = true;
      desktop.onUpdateEvent((next) => applyDesktopUpdateState(next));
    }
    try {
      const snapshot = await desktop.getUpdateState();
      if (snapshot) applyDesktopUpdateState(snapshot);
      if (!snapshot || snapshot.status === "idle") {
        lastUpdateCheckAt = Date.now();
        void desktop.checkUpdate?.().then((res) => applyDesktopUpdateState(res)).catch(() => {});
      }
    } catch {}
    return;
  }
  if (typeof desktop?.runAgent === "function") {
    void checkUpdateFallback({ silent: true });
  }
}

function maybeRefreshUpdateInBackground() {
  if (!state.me || Date.now() - lastUpdateCheckAt < 10 * 60 * 1000) return;
  const desktop = window.officeDesktop;
  if (typeof desktop?.checkUpdate === "function") {
    if (["downloading", "verifying", "installing", "ready"].includes(state.update?.status)) return;
    lastUpdateCheckAt = Date.now();
    void desktop.checkUpdate().then((res) => applyDesktopUpdateState(res)).catch(() => {});
  } else if (typeof desktop?.runAgent === "function") {
    void checkUpdateFallback({ silent: true });
  }
}

function homeRoute() {
  const role = state.me.role;
  return role === "admin" ? "admin" : role === "reviewer" ? "review" : "chat";
}

function routeParts() {
  return location.hash.slice(1).split("/").map((part) => {
    try { return decodeURIComponent(part); } catch { return ""; }
  });
}

function adminReset() {
  state.admin = { offset: 0, department: "", status: "", date: "" };
}

async function preloadSidebarData() {
  const me = state.me;
  if (!me) return;
  prefetchDesktopModels();
  try {
    const [conversations, todayData] = await Promise.all([
      api("/conversations").catch(() => null),
      api("/me/today").catch(() => null)
    ]);
    if (state.me !== me) return;
    if (Array.isArray(conversations)) {
      state.conversations = conversations;
      if (state.chatWorkspace) state.chatWorkspace.conversations = conversations;
      syncSidebarConversations();
      void autoTitleRecentConversations({ force: false, limit: 12 }).catch(() => {});
    }
    if (todayData?.items) {
      state.today = todayData.items;
      if (todayData.today) state.me.today = todayData.today;
      syncTodayBadge();
    }
  } catch {}
}

async function login(user) {
  await post("/auth/demo", { user_id: user });
  clearDrafts();
  resetChatState();
  sessionStorage.removeItem("qo-signed-out");
  clearAdmin();
  unmountShell();
  state.agent = null;
  agentDesktopModels = {};
  state.chatWorkspace = null;
  state.conversations = [];
  state.usageOpen = false;
  state.usageState = null;
  state.contextOpen = false;
  state.studioTab = "context";
  state.imageDrawerOpen = false;
  state._lastStudioKey = "";
  state.filesCache = [];
  state.filesFilter = "all";
  state.filesQuery = "";
  state.filter = "all";
  state.today = [];
  adminReset();
  historyOffset = 0;
  dialog.close();
  [state.me, state.engine] = await Promise.all([api("/me"), api("/engine")]);
  void preloadSidebarData();
  go(homeRoute());
}

async function logout() {
  await post("/auth/logout", {});
  clearDrafts();
  resetChatState();
  clearAdmin();
  unmountShell();
  sessionStorage.setItem("qo-signed-out", "1");
  state.me = null;
  state.agent = null;
  agentDesktopModels = {};
  state.chatWorkspace = null;
  state.conversations = [];
  state.usageOpen = false;
  state.usageState = null;
  state.filesCache = [];
  state.today = [];
  agentFiles = [];
  agentTask = null;
  dialog.close();
  history.replaceState(null, "", location.pathname);
  document.title = "登录 · 穹驭办公";
  root.innerHTML = loginView();
}

async function startup() {
  try {
    state.config = await api("/config");
    try { state.me = await api("/me"); }
    catch (error) {
      if (error.status !== 401) throw error;
      if (state.config.demo && !sessionStorage.getItem("qo-signed-out") &&
          new URLSearchParams(location.search).get("login") !== "1") {
        await login("lin");
        return;
      }
    }
    if (!state.me) {
      unmountShell();
      document.title = "登录 · 穹驭办公";
      root.innerHTML = loginView();
      return;
    }
    state.engine = await api("/engine");
    void preloadSidebarData();
    void initDesktopUpdater();
    if (!location.hash) go(homeRoute());
    else await render();
  } catch (error) {
    unmountShell();
    const switchServer = typeof window.officeDesktop?.openConnect === "function"
      ? button("切换工作台地址", "desktop-connect", { icon: "link" }) : "";
    root.innerHTML = `<section class="splash"><div class="splash-card">
      <div class="brand-mark">${icon("mark")}</div>
      <h1>无法连接工作台</h1>
      <p class="error">${e(error.message)}</p>
      <div class="actions">${button("重新连接", "boot", { cls: "primary", icon: "refresh" })}${switchServer}</div>
    </div></section>`;
  }
}

async function loadToday() {
  const data = await api("/me/today");
  state.me.today = data.today;
  state.today = data.items;
  syncTodayBadge();
}

function viewSignature() {
  return JSON.stringify([state.me.today, state.today.map((t) => [t.id, t.version, t.overdue])]);
}

function display(content, ticket, keepScroll = false) {
  if (ticket !== navigation) return;
  const y = window.scrollY;
  const workplace = ensureShell();
  if (!workplace) return;
  workplace.innerHTML = content;
  syncShell();
  if (!keepScroll) {
    triggerViewEnter();
    triggerTopbarSweep();
  }
  setMobileNav(false);
  syncShellCanvasState();
  window.scrollTo(0, keepScroll ? y : 0);
  if (state.route[0] === "chat") {
    syncDesktopUsagePanel();
    scrollToLatest();
    mountDeck();
    syncActiveChatView();
    if (!keepScroll && !state.chatWorkspace?.session?.turns?.length) {
      const textarea = document.querySelector("[data-form=conversation] textarea");
      if (textarea && !textarea.disabled) textarea.focus({ preventScroll: true });
    }
  }
}

async function fetchAgent(parts, force = false) {
  const [, id, kind] = parts;
  const taskID = id && id !== "new" ? id : "";
  const well = taskID + "|" + (kind || "");
  if (state.agent?.well !== well || force || !agentFiles.length) {
    const [task, files] = await Promise.all([taskID ? api("/assignments/" + encodeURIComponent(taskID)) : null, api("/files")]);
    if (task && (task.assignee_id !== state.me.id || !active(task.status))) throw new Error("该任务当前无法启动文稿协助");
    agentTask = task;
    agentFiles = files;
    if (state.agent?.well !== well) {
      state.agent = readDraft("agent/" + well) || {
        well, kind: kind === "pptx" || task?.kind === "pptx" ? "pptx" : "docx",
        title: (task?.title || "").slice(0, 110), instructions: "", pages: 8, material_ids: []
      };
    }
  }
  agentDesktopModels = await loadDesktopModelState({ allowPipeline: false, draftKey: "agent/last-model" });
  return agentSetup(state.agent, agentTask, agentFiles, agentDesktopModels);
}

async function render(keepScroll = false) {
  if (!state.me) return;
  const ticket = ++navigation;
  const parts = routeParts();
  const mode = views[parts[0]] ? parts[0] : homeRoute();
  const previous = currentMode;
  currentMode = mode;
  state.route = [mode, ...parts.slice(1)];
  if (["admin", "review"].includes(mode) && previous !== mode) adminReset();
  try {
    let content;
    switch (mode) {
      case "chat": {
        const workspace = await loadWorkspace(parts);
        if (ticket !== navigation) return;
        state.chatWorkspace = workspace;
        content = workspaceView(workspace);
        break;
      }
      case "today": await loadToday(); content = dashboard(); break;
      case "projects": content = await projectsView(parts[1]); break;
      case "task":
        state.focused = await api("/assignments/" + encodeURIComponent(parts[1] || ""));
        content = taskDetail(state.focused); break;
      case "history": {
        const list = await api("/me/tasks?offset=" + historyOffset + "&limit=" + historySize);
        content = historyView(list.items, historyOffset, historySize);
        break;
      }
      case "agent": content = await fetchAgent(state.route, true); break;
      case "files": {
        state.filesCache = await api("/files");
        content = filesView(state.filesCache);
        break;
      }
      case "settings":
        content = settingsView();
        if (!state.update || Date.now() - lastUpdateCheckAt > 3 * 60 * 1000) {
          void initDesktopUpdater();
        }
        break;
      case "admin":
        if (!isManager()) { go("today"); return; }
        content = await taskAdmin(); break;
      case "review":
        if (!isReview()) { go("today"); return; }
        content = await taskAdmin(true); break;
      case "audit":
        if (state.me.role !== "admin") { go("today"); return; }
        content = await auditPage(); break;
    }
    display(content, ticket, keepScroll);
  } catch (error) {
    if (ticket !== navigation) return;
    if (error.status === 401) {
      clearDrafts();
      unmountShell();
      state.me = null;
      document.title = "登录 · 穹驭办公";
      root.innerHTML = loginView("登录已过期，请重新登录。");
    } else if (keepScroll && !navigator.onLine && root.querySelector(".app-shell")) {
      syncConnectionIndicator();
    } else {
      display(pageHead("工作台请求失败") + `<div class="panel pad stack"><p class="error">${e(error.message)}</p><div class="actions">${button("重试", "refresh", { cls: "primary", icon: "refresh" })}${button("返回今日任务", "route", { id: "today" })}</div></div>`, ticket, keepScroll);
    }
  }
}

function localRefresh() {
  const workplace = root.querySelector("#workplace");
  if (!workplace) return;
  if (state.route[0] === "agent") workplace.innerHTML = agentSetup(state.agent, agentTask, agentFiles, agentDesktopModels);
  else if (state.route[0] === "today") workplace.innerHTML = dashboard();
  else if (state.route[0] === "files") {
    const listWrap = workplace.querySelector("[data-files-list]");
    if (listWrap) listWrap.innerHTML = renderFilesList(state.filesCache);
    for (const pill of workplace.querySelectorAll("[data-act=files-filter]")) {
      pill.classList.toggle("active", pill.dataset.id === state.filesFilter);
    }
  }
}

function saveAgent() {
  if (state.agent) saveDraft("agent/" + state.agent.well, state.agent);
}

function agentInput(form) {
  const data = new FormData(form);
  return {
    kind: data.get("kind"), title: data.get("title"), instructions: data.get("instructions"),
    pages: Number(data.get("pages") || 8), material_ids: data.getAll("material_ids")
  };
}

function planInput(form) {
  return {
    title: form.elements.plan_title.value, subtitle: form.elements.plan_subtitle.value,
    sections: [...form.querySelectorAll("[data-section]")].map((section) => {
      const fields = section.querySelectorAll("input,textarea");
      return { heading: fields[0].value, points: fields[1].value.split("\n").map((x) => x.trim()).filter(Boolean) };
    })
  };
}

function updateAgentRunStatus(form, message) {
  const status = form.querySelector("[data-agent-run-status]");
  if (!status) return;
  status.textContent = message;
  status.hidden = !message;
}

async function executeDesktopPlan(task, data, form) {
  if (typeof window.officeDesktop?.runAgent !== "function") throw new Error("请在桌面客户端中使用 AI 办理");
  const model = agentDesktopModels.selectedModel;
  if (!agentDesktopModels.models?.includes(model)) {
    throw new Error(agentDesktopModels.models?.length ? "请先选择可用模型" : agentDesktopModels.modelNotice || "暂无可用模型，请联系管理员开通。");
  }
  const base = task ? "/assignments/" + encodeURIComponent(task.id) + "/sessions" : "/agent/plan";
  let prepared;
  let phase = "prepare";
  try {
    prepared = await post(base + "/prepare", data);
    if (!prepared?.run_id) throw new Error("服务端未返回本轮运行编号");
    updateAgentRunStatus(form, "正在整理文稿…");
    phase = "agent";
    const completed = await runDesktopAgent({ system: prepared.system, prompt: prepared.prompt, context: prepared.context, model }, (label) => {
      if (typeof label === "string" && label) updateAgentRunStatus(form, label);
    }, (result) => {
      updateAgentRunStatus(form, "正在保存大纲…");
      phase = "complete";
      return post(base+"/complete",{
        run_id: prepared.run_id, text: result.text,
        usage: {
          ...(result.usage || {}), max_tokens_clamped: result.max_tokens_clamped,
          finish_reasons: result.finish_reasons, truncated: result.truncated
        }, model: result.model || model
      });
    });
    updateAgentRunStatus(form, "");
    return completed;
  } catch (error) {
    updateAgentRunStatus(form, "");
    if (prepared?.run_id) {
      try {
        await post(base + "/" + encodeURIComponent(prepared.run_id) + "/fail", {
          message: String(error.message || "桌面执行失败").slice(0, 1000)
        });
      } catch (failError) {
        if (phase === "complete" && failError.status === 409 && (!error.status || error.status >= 500)) {
          error.message = "无法确认本轮是否已保存，请刷新文稿页查看最新结果。";
        } else if (failError.status !== 409) {
          error.message += `（失败状态同步未成功：${failError.message}）`;
        }
      }
    }
    throw error;
  }
}

function submissionInput(form, save = false) {
  const data = new FormData(form);
  const fields = { note: data.get("note"), link: data.get("link"), artifact_ids: data.getAll("artifact_ids") };
  if (!form.dataset.key || save) form.dataset.key = uuid();
  saveDraft("submit/" + state.focused.id, { ...fields, key: form.dataset.key });
  return fields;
}

async function handleUpload(list, element) {
  if (!list.length) return;
  if (element?.closest("[data-form=conversation]")) return attachFiles(list, element);
  if (list.length > 12) throw new Error("请每次选择最多 12 个文件");
  const key = state.route[0] === "task" ? state.focused?.id : agentTask?.id && state.route[0] === "agent" ? agentTask.id : "";
  const mode = state.route[0];
  const label = element?.closest("[data-upload-zone]");
  label?.classList.add("drag");
  try {
    const created = await uploadSelected(list, key);
    if (mode === "agent" && state.route[0] === mode) {
      agentFiles = await api("/files");
      state.agent.material_ids = [...new Set([...state.agent.material_ids, ...created.filter((f) => f.extension !== "xlsx").map((f) => f.id)])].slice(0, 8);
      saveAgent();
      localRefresh();
    } else if (mode === "task" && key === state.focused?.id) {
      const draft = readDraft("submit/" + key) || {};
      saveDraft("submit/" + key, { ...draft, artifact_ids: [...new Set([...(draft.artifact_ids || []), ...created.map((x) => x.id)])], key: uuid() });
      await render(true);
    } else if (state.route[0] === mode) await render(true);
    toastMessage("文件已添加", "success");
  } finally {
    label?.classList.remove("drag");
    if (element?.matches("input")) element.value = "";
  }
}

const actionCopy = { extension: "申请延期", comment: "留言", revoke: "撤销任务", close: "关闭任务" };

async function onClick(event) {
  if (!event.target.closest("[data-form=conversation]")) closeSlashMenu();
  if (!event.target.closest(".composer-picker")) closeComposerPickers();
  if (event.target.closest("a[data-stop-prop]")) return;
  const target = event.target.closest("[data-act]");
  if (!target || target.disabled) return;
  const { act, id } = target.dataset;
  switch (act) {
    case "route":
      if (dialog.open) dialog.close();
      setMobileNav(false);
      go(id);
      break;
    case "toggle-sidebar": toggleSidebar(); break;
    case "toggle-theme": {
      const next = toggleTheme();
      syncThemeToggle();
      if (!state.me) root.innerHTML = loginView();
      else if (state.route[0] === "settings") {
        const workplace = root.querySelector("#workplace");
        if (workplace) workplace.innerHTML = settingsView();
      }
      toastMessage(next === "dark" ? "已切换至深色模式（白线）" : "已切换至浅色模式（黑线）");
      break;
    }
    case "open-mobile-nav": setMobileNav(true); break;
    case "close-mobile-nav": setMobileNav(false); break;
    case "toggle-context": toggleWorkspaceContext(target); break;
    case "studio-tab": switchStudioTab(id); break;
    case "deck-open": openDeckInStudio(id); break;
    case "boot": await startup(); break;
    case "open-task": dialog.close(); setMobileNav(false); go("task/" + encodeURIComponent(id)); break;
    case "agent-task": go("chat/task/" + encodeURIComponent(id)); break;
    case "chat-example": useExample(id); break;
    case "chat-generate-image": await generateImageFromComposer(target, render); break;
    case "toggle-image-drawer": toggleImageDrawer(); break;
    case "slash-trigger": triggerSlashMenu(); break;
    case "slash-pick": await applySlashCommand(id); break;
    case "picker-toggle": toggleComposerPicker(id); break;
    case "picker-pick": pickComposerOption(target.dataset.picker || "", id); break;
    case "image-lightbox": openImageLightbox(id, target.dataset.alt || "", target.dataset.meta || ""); break;
    case "chat-retry": retryTurn(id); break;
    case "chat-remove-source": removeSource(id); break;
    case "chat-sources": await showSources(); break;
    case "chat-history": showConversationHistory(); break;
    case "chat-auto-titles": await autoTitleRecentConversations({ force: true, limit: 15, triggerButton: target }); break;
    case "chat-earlier": await loadEarlier(); break;
    case "chat-cancel": await cancelActiveChat(); break;
    case "chat-resume": await resumeFailedTurn(target); break;
    case "chat-discard-checkpoint": await discardCheckpoint(); break;
    case "chat-abort-orphan": await abortOrphanTurn(id, target.dataset.kind || "turn", target); break;
    case "copy-reply": await copyReply(id, target); break;
    case "copy-invite": await copyText(id, "已复制项目邀请码"); break;
    case "desktop-connect": await window.officeDesktop?.openConnect?.(); break;
    case "usage-toggle": await toggleDesktopUsage(); break;
    case "usage-refresh": await refreshDesktopUsage(state.chatWorkspace); break;
    case "slides-clear": clearSelection(); break;
    case "project-create": projectDialog("create"); break;
    case "project-join": projectDialog("join"); break;
    case "project-rotate": await updateInvite(id, render); break;
    case "files-filter":
      state.filesFilter = id;
      localRefresh();
      break;
    case "profile":
      modal("账户与身份", `<div class="stack">
        <div class="profile-banner">
          <span class="avatar lg">${e(state.me.name.slice(-2))}</span>
          <div><h3>${e(state.me.name)}</h3><p>${e(state.me.department || "未分配部门")} · ${e(roles[state.me.role])}</p></div>
        </div>
        ${state.config.demo ? '<p class="hint">切换演示身份以体验员工、主管、审阅人与管理员视角。</p>' : ""}
        ${state.config.demo ? `<div class="demo-roster">${demoUsers.filter((u) => u[0] !== state.me.id).map(([person, name, role]) => `
          <button type="button" class="demo-user-card" data-act="demo-login" data-id="${e(person)}">
            <span class="avatar">${e(name.slice(-2))}</span>
            <span class="demo-user-meta"><strong>${e(name)}</strong><small>${e(role)}</small></span>
          </button>`).join("")}</div>` : ""}
        <div class="actions">${button("退出登录", "logout", { icon: "logout", cls: "danger" })}</div>
      </div>`);
      break;
    case "demo-login": await login(id); break;
    case "logout": await logout(); break;
    case "sso":
      if (!window.officeDesktop) { location.assign("/auth/start"); return; }
      target.disabled = true;
      target.textContent = "等待系统浏览器登录…";
      try {
        if (await window.officeDesktop.signIn()) await startup();
      } finally {
        target.disabled = false;
        target.textContent = "使用公司账号登录";
      }
      break;
    case "refresh": await render(true); break;
    case "filter-today": state.filter = id; localRefresh(); break;
    case "close-dialog": dialog.close(); break;
    case "history-prev": historyOffset = Math.max(0, historyOffset - historySize); await render(); break;
    case "history-next": historyOffset += historySize; await render(); break;
    case "audit-prev": state.auditOffset = Math.max(0, state.auditOffset - 20); await render(); break;
    case "audit-next": state.auditOffset += 20; await render(); break;
    case "admin-prev": state.admin.offset = Math.max(0, state.admin.offset - 20); await render(); break;
    case "admin-next": state.admin.offset += 20; await render(); break;
    case "admin-reset": adminReset(); await render(); break;
    case "new-task": await showTaskForm(); break;
    case "edit-plan": await showTaskForm(id); break;
    case "publish-plan": showPublish(id); break;
    case "change-task": await changeAssignment(state.focused); break;
    case "task-action": {
      state.focused = await post("/assignments/" + encodeURIComponent(state.focused.id) + "/action/" + id, { version: state.focused.version });
      await render(true);
      toastMessage(id === "start" ? "任务已开始办理" : "任务已暂停", "success");
      break;
    }
    case "task-note": {
      const message = id === "revoke" ? "撤销后员工停止办理，已提交文件和审阅记录仍可查看。" : id === "close" ? "关闭后停止办理，并保留历史交付。" : "";
      modal(actionCopy[id], `<form data-form="task-note" data-id="${id}" class="stack">
      ${message ? `<p class="notice warn">${message}</p>` : ""}
      <label class="field">说明<textarea name="note" rows="4" maxlength="3000" ${["extension", "comment"].includes(id) ? "required" : ""}></textarea></label>
      <div class="actions"><button class="button primary" type="submit">确认${actionCopy[id]}</button></div></form>`);
      break;
    }
    case "review-task":
    case "return-task": {
      const returned = act === "return-task";
      modal(returned ? "退回修改" : "验收成果", `<form class="stack" data-form="review" data-id="${e(id)}"
      data-decision="${returned ? "return" : "approve"}" data-key="${uuid()}">
      <label class="field">审核意见<textarea name="note" rows="4" maxlength="3000" ${returned ? "required" : ""} placeholder="${returned ? "说明需要调整的具体内容" : "填写通过说明（选填）"}"></textarea></label>
      <div class="actions"><button class="button ${returned ? "danger" : "primary"}" type="submit">${returned ? "确认退回" : "确认通过"}</button></div></form>`);
      break;
    }
    case "preview": {
      const file = await api("/files/" + encodeURIComponent(id) + "/preview");
      modal(file.name, `<div class="stack"><pre class="preview">${e(file.text || "此文件可下载后在相应应用中查看。")}</pre>
       <div class="actions"><a class="button primary" href="/api/v1/files/${encodeURIComponent(id)}/download">${icon("download")}下载文件</a>
       ${file.parent_id ? button("查看历史版本", "preview", { id: file.parent_id }) : ""}</div></div>`);
      break;
    }
    case "edit-file": {
      const file = await api("/files/" + encodeURIComponent(id) + "/preview");
      modal("修改办公文件", `<form class="stack" data-form="edit-file" data-id="${e(id)}">
       <p class="hint">在当前文件中替换同一段落内的连续文字，保存后自动保留旧版并生成新版本。</p>
       <pre class="preview">${e(file.text)}</pre>
       <label class="field">原文字<input required name="old" maxlength="300"></label>
       <label class="field">替换为<textarea name="new" rows="3" maxlength="500"></textarea></label>
       <div class="actions"><button class="button primary" type="submit">保存修改版</button></div></form>`);
      break;
    }
    case "choose-file": {
      const files = await api("/files");
      modal("关联已有办公文件", `<div class="file-list">${files.filter((f) => !f.assignment_id).map((f) => fileMarkup(f, { copy: true })).join("") || '<p class="hint">没有可关联的独立文稿。</p>'}</div>`);
      break;
    }
    case "copy-file": {
      const task = state.focused.id;
      const file = await post("/files/" + encodeURIComponent(id) + "/copy", { assignment_id: task });
      const draft = readDraft("submit/" + task) || {};
      saveDraft("submit/" + task, { ...draft, artifact_ids: [...new Set([...(draft.artifact_ids || []), file.id])], key: uuid() });
      dialog.close();
      await render(true);
      toastMessage("文件已关联到任务", "success");
      break;
    }
    case "manual-plan":
      if (!state.agent.plan) {
        state.agent.plan = { title: state.agent.title || "", subtitle: "", sections: [{ heading: "", points: [""] }] };
        state.agent.run_id = null;
        saveAgent();
        localRefresh();
      }
      document.querySelector("#plan input")?.focus();
      break;
    case "add-section": {
      const list = document.querySelector("#plan-sections");
      if (list.children.length >= 19) throw new Error("最多添加 19 节");
      list.insertAdjacentHTML("beforeend", planSection({ heading: "", points: [""] }, list.children.length));
      list.lastElementChild.querySelector("input").focus();
      state.agent.plan = planInput(list.closest("form"));
      saveAgent();
      break;
    }
    case "remove-section": {
      const section = target.closest("[data-section]");
      if (section.parentElement.children.length === 1) throw new Error("大纲至少保留一节");
      const form = section.closest("form");
      section.remove();
      [...form.querySelectorAll("[data-section] .plan-card-index")].forEach((node, idx) => {
        node.textContent = `第 ${String(idx + 1).padStart(2, "0")} 节`;
      });
      state.agent.plan = planInput(form);
      saveAgent();
      break;
    }
    case "check-engine": {
      state.engine = await api("/engine");
      if (state.engine.client_execution) {
        if (typeof window.officeDesktop?.models !== "function") {
          toastMessage("当前为浏览器访问，AI 办理请在穹驭办公桌面客户端中使用");
          break;
        }
        const info = await loadDesktopModelState({ force: true });
        if (info.models?.length) {
          const suffix = info.pipelineAvailable ? "（支持智能分工）" : "";
          toastMessage(`桌面连接正常，当前可用 ${info.models.length} 个模型${suffix}`, "success");
        } else {
          toastMessage(info.modelNotice || "暂无可用模型，请联系管理员开通", "error");
        }
      } else {
        toastMessage(state.engine.configured ? "模型配置已加载，可正常发起会话与起草" : "尚未配置模型连接", state.engine.configured ? "success" : "info");
      }
      break;
    }
    case "update-check": {
      if (typeof window.officeDesktop?.checkUpdate === "function") {
        const result = await window.officeDesktop.checkUpdate();
        applyDesktopUpdateState(result);
        if (result?.status === "up-to-date") toastMessage("当前已是最新版本", "success");
        else if (result?.status === "available") toastMessage(`发现新版本 v${result.latestVersion}`, "info");
        else if (result?.status === "error") toastMessage(result.error || "检查更新失败", "error");
      } else {
        await checkUpdateFallback({ silent: false });
      }
      break;
    }
    case "update-download": {
      if (typeof window.officeDesktop?.downloadUpdate !== "function") break;
      state.updateBannerDismissed = false;
      const result = await window.officeDesktop.downloadUpdate();
      applyDesktopUpdateState(result);
      if (result?.status === "ready") toastMessage(`v${result.latestVersion} 已下载完成，可立即安装并重启`, "success");
      else if (result?.status === "error") toastMessage(result.error || "下载更新失败", "error");
      break;
    }
    case "update-download-full": {
      if (typeof window.officeDesktop?.downloadUpdate !== "function") break;
      state.updateBannerDismissed = false;
      const result = await window.officeDesktop.downloadUpdate({ mode: "full" });
      applyDesktopUpdateState(result);
      if (result?.status === "ready") toastMessage(`v${result.latestVersion} 完整安装包已下载完成`, "success");
      else if (result?.status === "error") toastMessage(result.error || "下载更新失败", "error");
      break;
    }
    case "update-cancel": {
      if (typeof window.officeDesktop?.cancelUpdate === "function") {
        await window.officeDesktop.cancelUpdate();
        toastMessage("已取消下载");
      }
      break;
    }
    case "update-install": {
      if (typeof window.officeDesktop?.installUpdate !== "function") break;
      const result = await window.officeDesktop.installUpdate();
      applyDesktopUpdateState(result);
      break;
    }
    case "update-dismiss":
      state.updateBannerDismissed = true;
      syncUpdateUi();
      break;
    case "search":
      openCommandPalette();
      break;
  }
}

async function handleForm(event) {
  const form = event.target;
  if (!form.matches("form[data-form]")) return;
  event.preventDefault();
  if (form.dataset.busy) return;
  form.dataset.busy = "1";
  const controls = [...form.querySelectorAll("button[type=submit]")];
  controls.forEach((node) => { node.disabled = true; });
  try {
    const id = form.dataset.id;
    const type = form.dataset.form;
    const values = Object.fromEntries(new FormData(form));
    let message = "";
    if (type === "admin-filter") {
      state.admin = { status: values.status || "", department: values.department || "", date: values.date || "", offset: 0 };
      await render();
    } else if (type === "task-template") {
      const intent = event.submitter?.value || "draft";
      await saveTaskForm(form, intent);
      dialog.close();
      await render(true);
      message = intent === "publish" ? "任务已安排发布" : "任务已保存";
    } else if (type === "publish-plan") {
      await post("/tasks/" + encodeURIComponent(id) + "/publish", { version: Number(form.dataset.version) }, form.dataset.key);
      dialog.close();
      await render(true);
      message = "任务已安排发布";
    } else if (type === "change-task") {
      await api("/assignments/" + encodeURIComponent(state.focused.id), { method: "PATCH", body: assignmentValues(form, state.focused.version) });
      dialog.close();
      await render(true);
      message = "任务已更新";
    } else if (type === "task-note") {
      await post("/assignments/" + encodeURIComponent(state.focused.id) + "/action/" + id, { version: state.focused.version, note: values.note });
      dialog.close();
      await render(true);
      message = actionCopy[id] + "已记录";
    } else if (type === "submission") {
      const task = state.focused;
      const fields = submissionInput(form);
      await post("/assignments/" + encodeURIComponent(task.id) + "/submissions", { ...fields, version: task.version }, form.dataset.key);
      saveDraft("submit/" + task.id, null);
      await render(true);
      message = "成果已提交审核";
    } else if (type === "review") {
      await post("/submissions/" + encodeURIComponent(id) + "/review", { version: state.focused.version, decision: form.dataset.decision, note: values.note }, form.dataset.key);
      dialog.close();
      await render(true);
      message = form.dataset.decision === "approve" ? "成果已验收通过" : "已退回修改，员工可继续修订";
    } else if (type.startsWith("project-")) {
      const key = await submitProject(type, id, values);
      if (dialog.open) dialog.close();
      if (type === "project-context") {
        await render(true);
        message = "共享资料已发布";
      } else {
        go("projects/" + encodeURIComponent(key));
      }
    } else if (type === "conversation") {
      await sendMessage(form, render);
    } else if (type === "chat-sources") {
      chooseSources(form);
      dialog.close();
    } else if (type === "generate") {
      const data = agentInput(form);
      if (data.material_ids.length > 8) throw new Error("每次最多选择 8 份资料");
      const desktopRequired = state.engine?.client_execution === true;
      if (desktopRequired && typeof window.officeDesktop?.runAgent !== "function") {
        throw new Error("请在桌面客户端中使用 AI 办理");
      }
      if (desktopRequired && !agentDesktopModels.models?.includes(agentDesktopModels.selectedModel)) {
        throw new Error(agentDesktopModels.models?.length ? "请先选择可用模型" : agentDesktopModels.modelNotice || "暂无可用模型，请联系管理员开通。");
      }
      const task = agentTask;
      const agent = state.agent;
      const submit = controls[0];
      if (submit) submit.textContent = "正在起草…";
      Object.assign(agent, data);
      saveAgent();
      const result = desktopRequired ? await executeDesktopPlan(task, data, form) :
        await post(task ? "/assignments/" + encodeURIComponent(task.id) + "/sessions" : "/agent/plan", data);
      if (state.agent === agent && state.route[0] === "agent" && form.isConnected) {
        Object.assign(agent, { plan: result.plan, run_id: result.run_id, result: null, outputTelemetry: result.usage });
        saveAgent();
        localRefresh();
        document.querySelector("#plan")?.scrollIntoView({ behavior: "smooth", block: "nearest" });
      }
      message = "大纲已生成，可调整后保存文件";
    } else if (type === "render") {
      const plan = planInput(form);
      if (plan.sections.some((part) => part.points.length < 1 || part.points.length > 8)) throw new Error("每节正文请填写 1 至 8 行");
      const agent = state.agent;
      const task = agentTask;
      const file = await post("/agent/render", { kind: agent.kind, plan, assignment_id: task?.id || null, run_id: agent.run_id || null });
      if (state.agent === agent && state.route[0] === "agent" && form.isConnected) {
        Object.assign(agent, { plan, result: file });
        agentFiles = [file, ...agentFiles];
        saveAgent();
        if (task) {
          const draft = readDraft("submit/" + task.id) || {};
          saveDraft("submit/" + task.id, { ...draft, artifact_ids: [...new Set([...(draft.artifact_ids || []), file.id])], key: uuid() });
        }
        localRefresh();
      }
      message = "可编辑文件已生成并保存";
    } else if (type === "edit-file") {
      const changed = await post("/files/" + encodeURIComponent(id) + "/edit", { old: values.old, new: values.new });
      if (state.route[0] === "agent" && state.agent) {
        state.agent.result = changed;
        saveAgent();
      }
      if (state.route[0] === "task" && state.focused) {
        const task = state.focused.id;
        const draft = readDraft("submit/" + task) || {};
        const ids = (draft.artifact_ids || []).filter((f) => f !== id);
        saveDraft("submit/" + task, { ...draft, artifact_ids: [...new Set([...ids, changed.id])], key: uuid() });
      }
      dialog.close();
      await render(true);
      message = `已替换 ${changed.matches} 处并保存新版本`;
    }
    if (message) toastMessage(message, "success");
  } catch (error) {
    if (form.isConnected) errorIn(form, error);
  } finally {
    delete form.dataset.busy;
    if (form.dataset.form === "conversation") {
      if (form.isConnected) syncChatSubmit(form, state.chatWorkspace);
    } else {
      controls.forEach((node) => {
        if (form.dataset.form === "generate") {
          node.textContent = "AI 起草大纲";
          node.disabled = state.engine?.client_execution === true && (typeof window.officeDesktop?.runAgent !== "function" ||
            !agentDesktopModels.models?.includes(agentDesktopModels.selectedModel));
        } else node.disabled = false;
      });
    }
  }
}

document.addEventListener("click", (event) => {
  onClick(event).catch((error) => dialog.open ? errorIn(dialog.querySelector(".dialog-body"), error) : toastMessage(error.message, "error"));
});
document.addEventListener("submit", (event) => { void handleForm(event); });
document.addEventListener("input", (event) => {
  const element = event.target;
  if (element.matches("[data-files-search]")) {
    state.filesQuery = element.value;
    localRefresh();
    return;
  }
  if (element.matches("[data-form=conversation] textarea")) {
    syncComposerGrow(element);
    updateSlashMenu(element);
    saveWorkspaceInput(element.form);
  } else if (element.matches("[data-form=conversation] select")) {
    saveWorkspaceInput(element.form);
  }
  if (element.closest("form[data-form=submission]")) submissionInput(element.form, true);
  if (element.closest("form[data-form=generate]")) {
    const agent = state.agent;
    const data = agentInput(element.form);
    if (data.kind !== agent.kind) {
      agent.plan = null;
      agent.run_id = null;
      agent.result = null;
      root.querySelector("#plan")?.remove();
    }
    Object.assign(agent, data);
    const pages = element.form.querySelector("[data-pages]");
    if (pages) pages.hidden = agent.kind !== "pptx";
    saveAgent();
  }
  if (element.closest("form[data-form=render]")) {
    state.agent.plan = planInput(element.form);
    saveAgent();
  }
  if (element.name === "repeat" && element.closest("form[data-form=task-template]")) {
    for (const field of element.form.querySelectorAll("[data-repeat]")) field.hidden = element.value !== "weekdays";
  }
});
document.addEventListener("change", (event) => {
  if (event.target.matches("[name=allow_ai_images]")) {
    saveImageAutoPreference(event.target.checked);
    syncImageAutoIndicator();
    return;
  }
  if (event.target.matches("[data-chat-model]")) {
    selectChatModel(event.target);
    return;
  }
  if (event.target.matches("[data-agent-model]")) {
    selectDesktopModel(event.target, agentDesktopModels, "[data-agent-model-status]", "agent/last-model");
    return;
  }
  if (!event.target.matches("[data-upload]")) return;
  handleUpload([...event.target.files], event.target).catch((error) => toastMessage(error.message, "error"));
});
document.addEventListener("dragover", (event) => {
  const hasFiles = [...(event.dataTransfer?.types || [])].includes("Files");
  const zone = event.target.closest("[data-upload-zone]");
  if (zone) { event.preventDefault(); zone.classList.add("drag"); return; }
  const chatZone = hasFiles && event.target.closest("[data-chat-dropzone]");
  if (chatZone) { event.preventDefault(); chatZone.classList.add("drag-over"); }
});
document.addEventListener("dragleave", (event) => {
  event.target.closest("[data-upload-zone]")?.classList.remove("drag");
  const chatZone = event.target.closest("[data-chat-dropzone]");
  if (chatZone && !chatZone.contains(event.relatedTarget)) chatZone.classList.remove("drag-over");
});
document.addEventListener("drop", (event) => {
  const zone = event.target.closest("[data-upload-zone]");
  if (zone) {
    event.preventDefault();
    zone.classList.remove("drag");
    handleUpload([...event.dataTransfer.files], zone.querySelector("input")).catch((error) => toastMessage(error.message, "error"));
    return;
  }
  const chatZone = event.target.closest("[data-chat-dropzone]");
  if (chatZone && event.dataTransfer?.files?.length) {
    event.preventDefault();
    chatZone.classList.remove("drag-over");
    attachFiles([...event.dataTransfer.files]).catch((error) => toastMessage(error.message, "error"));
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    setMobileNav(false);
    closeSlashMenu();
    closeComposerPickers();
  }
  if (event.target.matches?.("[data-form=conversation] textarea")) {
    if (handleSlashKeydown(event)) return;
    if (event.key === "Enter" && !event.shiftKey && !event.isComposing && event.keyCode !== 229) {
      event.preventDefault();
      if (event.target.value.trim()) event.target.form.requestSubmit();
      return;
    }
  }
  if ((event.key === "Enter" || event.key === " ") && event.target.matches?.(".artifact-card[data-act=deck-open]")) {
    event.preventDefault();
    openDeckInStudio(event.target.dataset.id);
    return;
  }
  if (state.me && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
    event.preventDefault();
    openCommandPalette();
  }
});
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) maybeRefreshUpdateInBackground();
});
window.addEventListener("hashchange", () => { if (dialog.open) dialog.close(); void render(); });
window.addEventListener("qo:conversations-updated", () => { syncSidebarConversations(); });
for (const eventName of ["online", "offline"]) {
  window.addEventListener(eventName, () => {
    syncConnectionIndicator();
    if (state.route[0] === "today" && navigator.onLine) void render(true);
  });
}
setInterval(async () => {
  if (!state.me || document.hidden || dialog.open || state.route[0] !== "today" || !navigator.onLine || timerBusy) return;
  timerBusy = true;
  const me = state.me;
  try {
    const previous = viewSignature();
    await loadToday();
    if (state.me === me && state.route[0] === "today" && viewSignature() !== previous) {
      const workplace = root.querySelector("#workplace");
      if (workplace) workplace.innerHTML = dashboard();
    }
  } catch { /* Keep the last displayed task list during transient interruptions. */ }
  finally { timerBusy = false; }
}, 4000);

applyTheme(getTheme());
initCustomSelects();
void startup();
