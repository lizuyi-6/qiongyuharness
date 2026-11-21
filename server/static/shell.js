import { active, api, dialog, e, formatBytes, getTheme, go, icon, isManager, isReview, modal, roles, root, state } from "./ui.js";
import { renderUpdatePanelBody, views } from "./views.js";
import { conversationNav, runningConversationIds, syncDesktopUsagePanel } from "./chat.js";

const SIDEBAR_KEY = "qo-sidebar-collapsed";
let mountedUserId = null;
let lastConversationsSig = "";
let commandSearchTimer = null;
let commandTasksCache = null;

function readSidebarPreference() {
  try { return localStorage.getItem(SIDEBAR_KEY) === "1"; } catch { return false; }
}

function writeSidebarPreference(collapsed) {
  try { localStorage.setItem(SIDEBAR_KEY, collapsed ? "1" : "0"); } catch {}
}

export function brandMarkup() {
  return `<span class="brand-mark" aria-hidden="true">${icon("mark")}</span><span class="brand-name">穹驭办公</span>`;
}

function navItem(id) {
  const [route] = state.route;
  const on = route === id;
  const [name, glyph] = views[id];
  return `<button class="nav ${on ? "active" : ""}" type="button" data-act="route" data-id="${id}" data-nav-id="${id}" title="${name}" ${on ? 'aria-current="page"' : ""}>
    ${icon(glyph)}<span class="nav-text">${name}</span>${id === "today" ? '<span class="nav-count" data-today-count hidden></span>' : ""}${id === "settings" ? '<span class="nav-count update-badge" data-update-badge hidden>新</span>' : ""}${id === "feedback" ? '<span class="nav-count" data-feedback-count hidden></span>' : ""}
  </button>`;
}

export function isShellMounted() {
  return Boolean(root.querySelector(".app-shell") && mountedUserId && mountedUserId === state.me?.id);
}

export function unmountShell() {
  mountedUserId = null;
  lastConversationsSig = "";
  commandTasksCache = null;
}

export function ensureShell() {
  if (!state.me) return null;
  if (isShellMounted()) {
    syncShell();
    return root.querySelector("#workplace");
  }
  const person = state.me;
  const collapsed = readSidebarPreference();
  const offline = !navigator.onLine;
  const conversations = state.chatWorkspace?.conversations || state.conversations || [];
  mountedUserId = person.id;
  lastConversationsSig = "";

  root.innerHTML = `<div class="app-shell ${collapsed ? "sidebar-collapsed" : ""} ${state.route[0] === "chat" ? "workspace-route" : ""}">
    <div class="sidebar-backdrop" data-act="close-mobile-nav" aria-hidden="true"></div>
    <aside class="sidebar" aria-label="应用导航">
      <div class="sidebar-top">
        <div class="brand">${brandMarkup()}</div>
        <button type="button" class="icon-button sidebar-toggle" data-act="toggle-sidebar" aria-label="折叠或展开侧栏" title="折叠或展开侧栏">
          ${icon("sidebar")}
        </button>
      </div>
      <nav class="sidebar-nav">
        <div data-workspace-nav-slot>${conversationNav({ conversations })}</div>
        <div class="nav-group">
          <p class="nav-label">办公</p>
          ${["today", "projects", "files"].map(navItem).join("")}
        </div>
        <div class="nav-group subtle">
          ${["agent", "history"].map(navItem).join("")}
        </div>
        ${isReview() ? `<div class="nav-group">
          <p class="nav-label">管理</p>
          ${isManager() ? navItem("admin") : ""}${navItem("review")}${person.role === "admin" ? navItem("audit") : ""}
        </div>` : ""}
      </nav>
      <div class="sidebar-quota-slot" data-sidebar-quota hidden></div>
      <div class="sidebar-foot">
        ${navItem("feedback")}
        ${navItem("settings")}
        <button type="button" class="identity" data-act="profile" aria-label="账户与身份" title="${e(person.name)}">
          <span class="avatar">${e(person.name.slice(-2))}</span>
          <span class="id-copy">
            <strong>${e(person.name)}</strong>
            <small>${person.department ? `${e(person.department)} · ` : ""}${e(roles[person.role])}</small>
          </span>
          ${icon("down")}
        </button>
      </div>
    </aside>
    <div class="shell-main">
      <header class="topbar">
        <div class="topbar-left">
          <button type="button" class="icon-button mobile-menu-btn" data-act="open-mobile-nav" aria-label="打开导航菜单">
            ${icon("menu")}
          </button>
          <div class="crumbs" data-crumbs>
            <span>${e(person.department || "工作台")}</span>${icon("right")}<strong data-crumb-current>工作台</strong>
          </div>
        </div>
        <div class="top-actions">
          ${state.config?.demo ? '<button class="demo-mark" type="button" data-act="profile">演示环境</button>' : ""}
          <button class="search-button" type="button" data-act="search" aria-label="搜索任务或跳转页面">
            ${icon("search")}<span>搜索</span><kbd>Ctrl K</kbd>
          </button>
          <button class="icon-button bordered" type="button" data-act="feedback-open" aria-label="问题反馈与一键上传日志" title="问题反馈 / 一键上传日志">
            ${icon("bug")}
          </button>
          <button class="icon-button bordered" type="button" data-act="toggle-theme" data-theme-toggle aria-label="切换浅色或深色外观" title="切换浅色 / 深色">
            ${icon(getTheme() === "dark" ? "sun" : "moon")}
          </button>
          <span class="connection-state offline" data-connection-state ${offline ? "" : "hidden"}>
            <i class="dot"></i>离线
          </span>
        </div>
        <div class="topbar-line" data-topbar-line aria-hidden="true"></div>
      </header>
      <div class="update-banner" data-update-banner hidden></div>
      <main class="content" id="workplace"></main>
    </div>
  </div>`;

  syncShell();
  syncDesktopUsagePanel();
  return root.querySelector("#workplace");
}

export function syncShell() {
  const shellEl = root.querySelector(".app-shell");
  if (!shellEl || !state.me) return;

  const [route, sub] = state.route;
  shellEl.classList.toggle("workspace-route", route === "chat");

  const currentLabel = views[route]?.[0] || "工作台";
  const crumb = shellEl.querySelector("[data-crumb-current]");
  if (crumb) crumb.textContent = currentLabel;
  document.title = `${currentLabel} · 穹驭办公`;

  for (const btn of shellEl.querySelectorAll("[data-nav-id]")) {
    const on = btn.dataset.navId === route;
    btn.classList.toggle("active", on);
    if (on) btn.setAttribute("aria-current", "page");
    else btn.removeAttribute("aria-current");
  }

  syncSidebarConversations();
  syncTodayBadge();
  syncFeedbackBadge();
  syncConnectionIndicator();
  syncUpdateUi();
}

export function syncFeedbackBadge() {
  const badgeEl = root.querySelector("[data-feedback-count]");
  if (!badgeEl) return;
  const count = isManager() ? Number(state.feedbackOpenCount || 0) : 0;
  if (!count) {
    badgeEl.hidden = true;
    return;
  }
  badgeEl.hidden = false;
  badgeEl.textContent = count > 99 ? "99+" : String(count);
  badgeEl.classList.add("warn");
}

export function syncSidebarConversations() {
  const slot = root.querySelector("[data-workspace-nav-slot]");
  if (!slot) return;
  const conversations = state.chatWorkspace?.conversations || state.conversations || [];
  const activeId = state.route[0] === "chat" ? (state.route[1] || "") : "";
  const running = runningConversationIds();
  const sig = JSON.stringify([activeId, running, conversations.map((c) => [c.id, c.title, Boolean(c.pinned)])]);
  if (sig === lastConversationsSig) return;
  lastConversationsSig = sig;

  const prevScroll = slot.querySelector(".recent-sessions")?.scrollTop || 0;
  slot.innerHTML = conversationNav({ conversations });
  const nextList = slot.querySelector(".recent-sessions");
  if (nextList) nextList.scrollTop = prevScroll;
}

export function syncTodayBadge() {
  const badgeEl = root.querySelector("[data-today-count]");
  if (!badgeEl) return;
  const items = state.today || [];
  const pendingCount = items.filter((item) => active(item.status)).length;
  const hasAlert = items.some((item) => item.overdue || item.status === "returned");
  if (!pendingCount) {
    badgeEl.hidden = true;
    return;
  }
  badgeEl.hidden = false;
  badgeEl.textContent = pendingCount > 99 ? "99+" : String(pendingCount);
  badgeEl.classList.toggle("warn", hasAlert);
}

export function syncConnectionIndicator() {
  const indicator = root.querySelector("[data-connection-state]");
  if (!indicator) return;
  indicator.hidden = navigator.onLine;
}

export function syncUpdateUi() {
  const info = state.update;
  const nativeUpdater = typeof window.officeDesktop?.checkUpdate === "function";
  const hasActionable = Boolean(
    info && (
      ["available", "downloading", "verifying", "ready", "installing"].includes(info.status) ||
      (!nativeUpdater && info.hasUpdate && info.full?.url)
    )
  );

  const navBadge = root.querySelector("[data-update-badge]");
  if (navBadge) {
    navBadge.hidden = !hasActionable;
    navBadge.textContent = info?.status === "ready" ? "待装" : "新";
  }

  const panel = root.querySelector("[data-update-panel]");
  if (panel) {
    panel.innerHTML = renderUpdatePanelBody();
  }

  const banner = root.querySelector("[data-update-banner]");
  if (!banner) return;
  if (!hasActionable || state.updateBannerDismissed) {
    banner.hidden = true;
    banner.innerHTML = "";
    return;
  }

  const rawVer = String(info?.latestVersion || "").trim();
  const ver = rawVer ? (/^v/i.test(rawVer) ? rawVer : `v${rawVer}`) : "";
  const isDelta = info?.mode === "delta";

  if (!nativeUpdater && info?.full?.url) {
    banner.hidden = false;
    banner.innerHTML = `<div class="update-banner-inner">
      <div class="update-banner-copy">
        <span class="update-banner-dot" aria-hidden="true"></span>
        <strong>发现新版本 ${e(ver)}</strong>
        <span>升级后支持应用内增量更新（${e(formatBytes(info.full.size))}）</span>
      </div>
      <div class="update-banner-actions">
        <a class="button sm primary" href="${e(info.full.url)}" download>${icon("download")}下载安装包</a>
        <button type="button" class="icon-button sm" data-act="update-dismiss" aria-label="稍后提醒" title="稍后提醒">${icon("x")}</button>
      </div>
    </div>`;
    return;
  }

  if (info.status === "downloading" || info.status === "verifying") {
    const pct = Math.max(0, Math.min(100, Number(info.percent) || 0));
    const speed = info.bytesPerSecond > 0 ? ` · ${formatBytes(info.bytesPerSecond)}/s` : "";
    banner.hidden = false;
    banner.innerHTML = `<div class="update-banner-inner">
      <div class="update-banner-copy">
        <span class="update-banner-dot is-pulse" aria-hidden="true"></span>
        <strong>${info.status === "verifying" ? `正在校验 ${e(ver)}` : `正在下载 ${e(ver)}（${pct}%）`}</strong>
        <span>${info.status === "verifying" ? "校验文件完整性…" : `${e(isDelta ? "增量更新" : "完整包")}${e(speed)}`}</span>
        <span class="update-banner-bar"><i style="width:${pct}%"></i></span>
      </div>
      <div class="update-banner-actions">
        ${info.status === "downloading" ? '<button type="button" class="text-button" data-act="update-cancel">取消</button>' : ""}
      </div>
    </div>`;
    return;
  }

  if (info.status === "ready") {
    banner.hidden = false;
    banner.innerHTML = `<div class="update-banner-inner">
      <div class="update-banner-copy">
        <span class="update-banner-dot is-ready" aria-hidden="true"></span>
        <strong>${e(ver)} 已下载完成</strong>
        <span>点击即可安装并自动重启，会话与文件不会丢失</span>
      </div>
      <div class="update-banner-actions">
        <button type="button" class="button sm primary" data-act="update-install">${icon("refresh")}立即安装并重启</button>
        <button type="button" class="icon-button sm" data-act="update-dismiss" aria-label="稍后安装" title="稍后安装">${icon("x")}</button>
      </div>
    </div>`;
    return;
  }

  if (info.status === "installing") {
    banner.hidden = false;
    banner.innerHTML = `<div class="update-banner-inner">
      <div class="update-banner-copy">
        <span class="update-banner-dot is-pulse" aria-hidden="true"></span>
        <strong>正在应用 ${e(ver)} 并重启…</strong>
      </div>
    </div>`;
    return;
  }

  banner.hidden = false;
  banner.innerHTML = `<div class="update-banner-inner">
    <div class="update-banner-copy">
      <span class="update-banner-dot" aria-hidden="true"></span>
      <strong>发现新版本 ${e(ver)}</strong>
      <span>${e(isDelta ? `增量补丁仅 ${formatBytes(info.size)}` : `完整包 ${formatBytes(info.size)}`)}</span>
    </div>
    <div class="update-banner-actions">
      <button type="button" class="button sm primary" data-act="update-download">${icon("download")}立即更新</button>
      <button type="button" class="text-button" data-act="route" data-id="settings">详情</button>
      <button type="button" class="icon-button sm" data-act="update-dismiss" aria-label="稍后提醒" title="稍后提醒">${icon("x")}</button>
    </div>
  </div>`;
}

export function triggerViewEnter() {
  const workplace = root.querySelector("#workplace");
  if (!workplace) return;
  workplace.classList.remove("view-entering");
  void workplace.offsetWidth;
  workplace.classList.add("view-entering");
}

export function syncThemeToggle() {
  const dark = getTheme() === "dark";
  for (const btn of root.querySelectorAll("[data-theme-toggle]")) {
    btn.innerHTML = icon(dark ? "sun" : "moon");
    btn.title = dark ? "切换到浅色（黑线）" : "切换到深色（白线）";
  }
}

export function triggerTopbarSweep() {
  const line = root.querySelector("[data-topbar-line]");
  if (!line) return;
  line.classList.remove("sweeping");
  void line.offsetWidth;
  line.classList.add("sweeping");
}

export function toggleSidebar() {
  const shellEl = root.querySelector(".app-shell");
  if (!shellEl) return;
  const collapsed = shellEl.classList.toggle("sidebar-collapsed");
  writeSidebarPreference(collapsed);
}

export function setMobileNav(open) {
  const shellEl = root.querySelector(".app-shell");
  if (!shellEl) return;
  shellEl.classList.toggle("mobile-nav-open", Boolean(open));
}

function quickCommands() {
  const list = [
    { label: "新建办公会话", hint: "AI 协助与文件生成", icon: "plus", act: "route", id: "chat" },
    { label: "今日任务", hint: "查看待办与截止提醒", icon: "today", act: "route", id: "today" },
    { label: "新建文稿", hint: "Word / PPT 大纲工坊", icon: "docx", act: "route", id: "agent" },
    { label: "我的项目", hint: "项目协作与共享资料", icon: "layers", act: "route", id: "projects" },
    { label: "我的文件", hint: "预览、修改与下载文件", icon: "folder", act: "route", id: "files" },
    { label: "任务历史", hint: "历史分派记录", icon: "history", act: "route", id: "history" },
    { label: "问题反馈与日志", hint: "上报问题与查看诊断日志", icon: "bug", act: "route", id: "feedback" },
    { label: "一键上传日志反馈", hint: "快速打包桌面与前端日志", icon: "upload", act: "feedback-open", id: "" },
    { label: "工作台设置", hint: "身份与模型连接状态", icon: "sliders", act: "route", id: "settings" }
  ];
  if (isManager()) {
    list.splice(2, 0,
      { label: "新建并分派任务", hint: "管理员操作", icon: "plus", act: "new-task", id: "" },
      { label: "任务管理", hint: "全局任务与重复计划", icon: "tasks", act: "route", id: "admin" }
    );
  }
  if (isReview()) {
    list.push({ label: "审核队列", hint: "待验收成果", icon: "check", act: "route", id: "review" });
  }
  if (state.me?.role === "admin") {
    list.push({ label: "审计记录", hint: "系统与模型操作日志", icon: "audit", act: "route", id: "audit" });
  }
  return list;
}

function renderCommandResults(container, query, tasks = null) {
  const q = query.trim().toLocaleLowerCase();
  const commands = quickCommands().filter((item) =>
    !q || item.label.toLocaleLowerCase().includes(q) || item.hint.toLocaleLowerCase().includes(q)
  );
  const matchedTasks = q && Array.isArray(tasks)
    ? tasks.filter((t) => t.title.toLocaleLowerCase().includes(q)).slice(0, 12)
    : [];

  let html = "";
  if (commands.length) {
    html += `<p class="command-section-label">${q ? "匹配功能" : "快捷跳转"}</p><div class="command-list">${commands.map((item, idx) => `
      <button type="button" class="command-item ${idx === 0 && !matchedTasks.length ? "active" : ""}" data-cmd-item data-act="${e(item.act)}" data-id="${e(item.id)}">
        ${icon(item.icon)}<span>${e(item.label)}</span><small>${e(item.hint)}</small>
      </button>`).join("")}</div>`;
  }
  if (q) {
    html += `<p class="command-section-label">匹配任务</p>`;
    if (tasks === null) {
      html += `<p class="hint">正在检索最近任务…</p>`;
    } else if (matchedTasks.length) {
      html += `<div class="command-list">${matchedTasks.map((task, idx) => `
        <button type="button" class="command-item ${!commands.length && idx === 0 ? "active" : ""}" data-cmd-item data-act="open-task" data-id="${e(task.id)}">
          ${icon(task.kind || "docx")}<span>${e(task.title)}</span><small>${e(task.work_date || "")}</small>
        </button>`).join("")}</div>`;
    } else {
      html += `<p class="hint">没有找到包含「${e(query.trim())}」的近期任务。</p>`;
    }
  }
  container.innerHTML = html || `<p class="hint">没有匹配的操作或任务。</p>`;
}

export function openCommandPalette() {
  modal("快捷搜索与导航", `<div class="command-box" data-command-palette>
    <div class="command-input-wrap">
      ${icon("search")}
      <input type="search" name="query" maxlength="130" placeholder="输入任务标题关键词，或搜索功能页面…" autocomplete="off" data-command-input>
      <kbd>ESC</kbd>
    </div>
    <div class="stack-sm" data-command-results></div>
  </div>`);

  const input = dialog.querySelector("[data-command-input]");
  const results = dialog.querySelector("[data-command-results]");
  if (!input || !results) return;

  renderCommandResults(results, "", commandTasksCache);
  input.focus();

  const loadTasks = async (query) => {
    try {
      if (!commandTasksCache) {
        const endpoint = isReview() ? "/admin/tasks?limit=200" : "/me/tasks?limit=200";
        const data = await api(endpoint);
        commandTasksCache = data.items || data;
      }
      if (dialog.open && input.isConnected) {
        renderCommandResults(results, input.value, commandTasksCache);
      }
    } catch {
      if (dialog.open && input.isConnected && query.trim()) {
        results.insertAdjacentHTML("beforeend", '<p class="error">任务搜索暂时不可用</p>');
      }
    }
  };

  input.addEventListener("input", () => {
    const val = input.value;
    renderCommandResults(results, val, commandTasksCache);
    clearTimeout(commandSearchTimer);
    if (val.trim() && !commandTasksCache) {
      commandSearchTimer = setTimeout(() => { void loadTasks(val); }, 180);
    }
  });

  input.addEventListener("keydown", (event) => {
    const items = [...results.querySelectorAll("[data-cmd-item]")];
    if (!items.length) return;
    const currentIndex = items.findIndex((item) => item.classList.contains("active"));
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      const delta = event.key === "ArrowDown" ? 1 : -1;
      const next = (currentIndex + delta + items.length) % items.length;
      items.forEach((el, idx) => el.classList.toggle("active", idx === next));
      items[next].scrollIntoView({ block: "nearest" });
    } else if (event.key === "Enter") {
      event.preventDefault();
      const target = items[currentIndex >= 0 ? currentIndex : 0];
      if (target) {
        const { act, id } = target.dataset;
        dialog.close();
        if (act === "route") go(id);
        else if (act === "open-task") go("task/" + encodeURIComponent(id));
        else target.click();
      }
    }
  });
}
