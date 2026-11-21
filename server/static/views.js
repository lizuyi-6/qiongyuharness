import { active, badge, button, dateLabel, deadline, e, eventLabels, fileIcon, fileMarkup, formatBytes, getTheme,
  icon, isManager, isReview, labels, outputWarnings, priorities, readDraft, roles, state, uploadMarkup } from "./ui.js";

export const views = {
  today: ["今日任务", "today"],
  agent: ["文稿", "docx"],
  files: ["我的文件", "folder"],
  chat: ["办公会话", "message"],
  projects: ["项目", "layers"],
  history: ["任务历史", "history"],
  admin: ["任务管理", "tasks"],
  review: ["审核队列", "check"],
  audit: ["审计记录", "audit"],
  feedback: ["问题反馈", "bug"],
  settings: ["设置", "sliders"],
  task: ["任务详情", "docx"]
};

export const demoUsers = [
  ["lin", "林知夏", "员工 · 产品与运营"],
  ["admin", "李祖祎", "超级管理员"],
  ["manager", "周婧", "主管 · 产品与运营"],
  ["chen", "陈以宁", "员工 · 产品与运营"],
  ["song", "宋予安", "员工 · 技术研发"],
  ["reviewer", "顾言", "审阅人 · 产品与运营"]
];

function brandLockup() {
  return `<div class="brand"><span class="brand-mark" aria-hidden="true">${icon("mark")}</span><span class="brand-name">穹驭办公</span></div>`;
}

export function loginView(error = "") {
  const isDemo = Boolean(state.config?.demo);
  const dark = getTheme() === "dark";
  const switchServer = typeof window.officeDesktop?.openConnect === "function"
    ? `<button type="button" class="text-button" data-act="desktop-connect">${icon("link")}切换工作台服务地址</button>`
    : "<span></span>";
  return `<div class="login"><section class="login-card ${isDemo ? "demo-wide" : ""}">
    <div class="login-head">
      ${brandLockup()}
      <h1>${isDemo ? "选择演示身份" : "登录穹驭办公"}</h1>
    </div>
    ${error ? `<p class="error" role="alert">${e(error)}</p>` : ""}
    ${isDemo
      ? `<div class="demo-roster">${demoUsers.map(([id, name, role]) => `
          <button type="button" class="demo-user-card ${id === "lin" ? "featured" : ""}" data-act="demo-login" data-id="${e(id)}">
            <span class="avatar">${e(name.slice(-2))}</span>
            <span class="demo-user-meta"><strong>${e(name)}</strong><small>${e(role)}</small></span>
          </button>`).join("")}</div>`
      : state.config?.sso
        ? `<div class="stack-sm">${button("使用公司账号登录", "sso", { cls: "primary" })}</div>`
        : '<p class="notice warn">企业登录入口尚未配置。</p>'}
    <div class="login-foot">${switchServer}<button type="button" class="text-button" data-act="toggle-theme">${icon(dark ? "sun" : "moon")}${dark ? "浅色模式" : "深色模式"}</button></div>
  </section></div>`;
}

export function pageHead(title, description = "", action = "") {
  return `<header class="page-head"><div class="page-head-copy"><h1>${e(title)}</h1>${description ? `<p>${e(description)}</p>` : ""}</div>${action}</header>`;
}

function taskTable(items, emptyMessage) {
  return `<div class="table-wrap"><table class="work-list">
    <thead><tr><th scope="col">任务</th><th scope="col">状态</th><th scope="col">截止时间</th></tr></thead>
    <tbody>${items.length ? items.map((task) => `<tr><td>
      <button class="task-link" type="button" data-act="open-task" data-id="${e(task.id)}" aria-label="${e(task.title)}">
        ${fileIcon(task.kind)}
        <span class="task-link-body">
          <strong>${e(task.title)}</strong>
          <small>${e(labels[task.kind] || task.kind)}${task.priority !== "normal" ? ` · <span class="${task.priority === "urgent" ? "late" : ""}">${e(priorities[task.priority])}</span>` : ""}</small>
        </span>
      </button>
    </td><td>${badge(task.status)}</td><td class="${task.overdue ? "late" : ""}">
      ${e(deadline(task.due_at))}${task.overdue ? '<span class="deadline-note">逾期</span>' : ""}
    </td></tr>`).join("") : `<tr><td colspan="3"><p class="empty">${e(emptyMessage)}</p></td></tr>`}</tbody>
  </table></div>`;
}

function taskSubset(items, key) {
  if (key === "working") return items.filter((x) => ["pending", "in_progress", "paused"].includes(x.status));
  if (key === "done") return items.filter((x) => ["submitted", "approved"].includes(x.status));
  if (key === "attention") return items.filter((x) => x.overdue || x.status === "returned");
  return items;
}

export function dashboard() {
  const { today, filter } = state;
  const ordered = [...today].sort((a, b) => {
    if (Boolean(a.overdue) !== Boolean(b.overdue)) return Number(Boolean(b.overdue)) - Number(Boolean(a.overdue));
    if (Boolean(a.due_at) !== Boolean(b.due_at)) return Number(Boolean(b.due_at)) - Number(Boolean(a.due_at));
    return (a.due_at || "").localeCompare(b.due_at || "");
  });
  const reminders = ordered.filter((task) => active(task.status) && task.due_at).slice(0, 4);
  const tabs = [["all", "全部"], ["working", "待办理"], ["done", "已提交／验收"], ["attention", "退回／逾期"]];
  const date = dateLabel(state.me.today, { month: "long", day: "numeric", weekday: "long" });

  return `${pageHead("今日任务", "", `<time class="head-date" datetime="${e(state.me.today)}">${icon("calendar")}${e(date)}</time>`)}
  <div class="dashboard">
    <section class="daily-list" aria-label="任务列表">
      <div class="tabs" aria-label="任务状态">${tabs.map(([key, title]) => `
        <button class="tab ${filter === key ? "active" : ""}" type="button" data-act="filter-today" data-id="${key}" aria-pressed="${filter === key}">
          ${title}<span class="tab-count">${taskSubset(today, key).length}</span>
        </button>`).join("")}</div>
      ${taskTable(taskSubset(ordered, filter), "当前分类暂无任务")}
    </section>
    <aside class="deadlines" aria-labelledby="deadlines-title">
      <h2 id="deadlines-title">截止提醒</h2>
      ${reminders.length ? `<ul class="agenda">${reminders.map((task) => `<li>
        <button class="agenda-item" type="button" data-act="open-task" data-id="${e(task.id)}" aria-label="查看截止任务：${e(task.title)}">
          <strong>${e(task.title)}</strong>
          <span class="${task.overdue ? "late" : ""}">${e(deadline(task.due_at))}${task.overdue ? " · 已逾期" : ""}</span>
        </button></li>`).join("")}</ul>` : '<p class="hint">暂无即将到期的任务</p>'}
    </aside>
  </div>`;
}

export function historyView(items, offset, pageSize) {
  return `${pageHead("任务历史")}
  <section aria-label="已发布任务">
    ${taskTable(items, "暂无历史任务")}
    ${pagination("history", offset, pageSize, items.length)}
  </section>`;
}

export function pagination(prefix, offset, size, found) {
  return `<div class="pagination"><span class="hint">第 ${Math.floor(offset / size) + 1} 页</span>
  <div class="pagination-controls">${button("上一页", prefix + "-prev", { disabled: offset === 0, cls: "sm" })}${button("下一页", prefix + "-next", { disabled: found < size, cls: "sm" })}</div></div>`;
}

function timeline(task) {
  return task.events.map((event) => {
    const changeNames = { title: "标题", description: "要求", deliverable: "交付标准", due_at: "截止时间", priority: "优先级", reviewer_id: "审核人", assignee_id: "员工" };
    const changes = event.type === "task.updated" ? Object.keys(event.payload.after || {}).map((key) => {
      const before = event.payload.before[key];
      const after = event.payload.after[key];
      if (["reviewer_id", "assignee_id"].includes(key)) return `${changeNames[key]}已调整`;
      const value = (v) => key === "due_at" ? deadline(v) : key === "priority" ? priorities[v] : String(v || "—");
      return `${changeNames[key] || key}：${value(before)} → ${value(after)}`;
    }).join("\n") : "";
    return `<div><strong>${e(eventLabels[event.type] || event.type)}</strong>
    <small>${e(event.name || "系统")} · ${e(deadline(event.created_at))}</small>
    ${event.payload.note || changes ? `<p class="prose hint">${e(event.payload.note || changes)}</p>` : ""}</div>`;
  }).join("") || '<p class="hint">暂无办理记录</p>';
}

function reviewBox(t, submission, canReview) {
  const showReview = canReview && t.status === "submitted" && submission.status === "submitted";
  return `<div class="review-box"><div class="between"><strong>第 ${submission.revision} 版</strong>${badge(submission.status)}</div>
    <p class="hint">${e(deadline(submission.created_at))}</p>
    ${submission.note ? `<p class="prose">${e(submission.note)}</p>` : ""}
    ${submission.link ? `<a href="${e(submission.link)}" target="_blank" rel="noopener noreferrer">${icon("link")} 打开交付链接</a>` : ""}
    ${submission.files.length ? `<div class="file-list">${submission.files.map((f) => fileMarkup(f)).join("")}</div>` : ""}
    ${submission.review_note ? `<p class="notice ${submission.status === "returned" ? "warn" : ""}">${e(submission.review_note)}</p>` : ""}
    ${showReview ? `<div class="actions">${button("通过验收", "review-task", { cls: "primary", id: submission.id })}
      ${button("退回修改", "return-task", { cls: "danger", id: submission.id })}</div>` : ""}
  </div>`;
}

export function taskDetail(t) {
  const mine = state.me.id === t.assignee_id;
  const canManage = isManager();
  const canReview = !mine && (canManage || state.me.id === t.reviewer_id);
  const canDo = mine && active(t.status);
  const note = readDraft("submit/" + t.id) || {};
  const picked = new Set(note.artifact_ids || []);
  const late = t.overdue ? badge("overdue") : "";
  return `${pageHead(t.title, "由 " + t.owner_name + " 分派给 " + t.assignee_name,
    `<div class="actions">${button("返回", "route", { id: mine ? "today" : isManager() ? "admin" : "review", icon: "left" })}
    ${canManage && !["approved", "revoked", "closed"].includes(t.status) ? button("调整任务", "change-task", { icon: "edit" }) : ""}</div>`)}
  <div class="columns"><div class="stack">
    <section class="panel pad stack">
      <div class="summary">
        <div><span class="hint">任务状态</span><div class="actions">${badge(t.status)} ${late}</div></div>
        <div><span class="hint">截止时间</span><strong class="${t.overdue ? "late" : ""}">${e(deadline(t.due_at))}</strong></div>
        <div><span class="hint">优先级</span><strong>${e(priorities[t.priority])}</strong></div>
      </div>
      <section><h2 class="group-title">任务要求</h2><p class="prose">${e(t.description || "暂无补充要求")}</p></section>
      <section><h2 class="group-title">交付标准</h2><p class="prose">${e(t.deliverable || "由经办员工与主管确认")}</p></section>
      ${t.materials.length ? `<section><h2 class="group-title">任务资料</h2><div class="file-list">${t.materials.map((f) => fileMarkup(f)).join("")}</div></section>` : ""}
      ${canDo ? `<div class="actions">${t.status === "in_progress" ? button("暂停办理", "task-action", { id: "pause" }) : button("开始办理", "task-action", { id: "start", cls: "primary" })}
        ${button("申请延期", "task-note", { id: "extension" })}${button("留言", "task-note", { id: "comment" })}</div>` : ""}
      ${canManage && !["approved", "revoked", "closed"].includes(t.status) ? `<div class="actions">
        ${button("撤销任务", "task-note", { id: "revoke", cls: "danger" })}${t.status !== "scheduled" ? button("手工关闭", "task-note", { id: "close" }) : ""}</div>` : ""}
    </section>
    ${canDo ? `<section class="panel pad stack"><h2 class="section-title">提交成果</h2>
      <div class="actions">${uploadMarkup("上传成果")}
      ${button("从我的文件关联", "choose-file", { icon: "folder" })}
      ${button("用会话办理", "agent-task", { id: t.id, icon: "message" })}</div>
      <form data-form="submission" class="stack" data-key="${e(note.key || "")}">
        ${t.files.length ? `<div class="file-list">${t.files.map((f) => fileMarkup(f, { check: "artifact_ids", checked: picked.has(f.id), edit: true })).join("")}</div>` : ""}
        <label class="field">成果说明<textarea name="note" rows="3" maxlength="4000" placeholder="简述本次交付内容和需要审核的事项">${e(note.note || "")}</textarea></label>
        <label class="field">交付链接（选填）<input name="link" type="url" maxlength="500" value="${e(note.link || "")}" placeholder="https://"></label>
        <div class="actions"><button class="button primary" type="submit">${icon("check")}提交审核</button></div>
      </form></section>` : ""}
    <section class="panel pad stack"><h2 class="section-title">历史交付</h2>
      ${t.submissions.map((sub) => reviewBox(t, sub, canReview)).join("") || '<p class="hint">暂无提交记录</p>'}</section>
  </div><section class="panel pad stack sticky-col"><h2 class="section-title">办理记录</h2><div class="timeline">${timeline(t)}</div></section></div>`;
}

export function agentSetup(session, task, files, desktopModels = {}) {
  const plan = session.plan;
  const desktopRequired = state.engine?.client_execution === true;
  const desktopAvailable = typeof window.officeDesktop?.runAgent === "function";
  const modelSelected = Boolean(desktopModels.selectedModel && desktopModels.models?.includes(desktopModels.selectedModel));
  const blocked = desktopRequired && !desktopAvailable;
  const modelRequired = desktopRequired && desktopAvailable;
  const saved = [...(task?.materials || []), ...files.filter((f) => !f.assignment_id || f.assignment_id === task?.id)]
    .filter((f) => ["docx", "pptx", "pdf", "txt", "md"].includes(f.extension)).slice(0, 20);

  return `${pageHead("文稿", task ? "关联任务 · " + task.title : "",
    task ? button("返回任务", "open-task", { id: task.id, icon: "left" }) : "")}
    <div class="columns"><div class="stack"><section class="panel pad stack">
      <h2 class="section-title">文稿信息</h2>
      <form class="stack" data-form="generate">
        <div class="form-grid">
          <label class="field span">标题<input required name="title" maxlength="110" value="${e(session.title)}" placeholder="例如：产品方案汇报"></label>
          <label class="field">输出格式<select name="kind"><option value="docx" ${session.kind === "docx" ? "selected" : ""}>Word（DOCX）</option><option value="pptx" ${session.kind === "pptx" ? "selected" : ""}>PowerPoint（PPTX）</option></select></label>
          <label class="field" data-pages ${session.kind !== "pptx" ? "hidden" : ""}>页数（含封面）<input name="pages" type="number" min="3" max="20" value="${e(session.pages || 8)}"></label>
          ${modelRequired ? `<label class="field span agent-model-picker">模型<select data-agent-model aria-label="执行模型"
            ${desktopModels.models?.length ? "" : "disabled"}>
            <option value="">${desktopModels.models?.length ? "请选择模型" : "暂无可用模型"}</option>
            ${(desktopModels.models || []).map((model) => `<option value="${e(model)}" ${desktopModels.selectedModel === model ? "selected" : ""}>${e(model)}</option>`).join("")}
          </select></label>` : ""}
          <label class="field span">写作要求<textarea name="instructions" rows="3" maxlength="6000" placeholder="受众、重点和格式要求">${e(session.instructions || "")}</textarea></label>
        </div>
        <div class="draft-sources">
          ${saved.length ? `<details ${session.material_ids.length ? "open" : ""}><summary>选择参考资料${session.material_ids.length ? `（已选 ${session.material_ids.length} 份）` : ""}</summary>
            <div class="file-list">${saved.map((f) => fileMarkup(f, { check: "material_ids", checked: session.material_ids.includes(f.id) })).join("")}</div></details>` : ""}
          ${uploadMarkup("上传参考资料")}
        </div>
        ${blocked ? '<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>' : ""}
        ${modelRequired && desktopModels.modelNotice ? `<p class="agent-model-status ${desktopModels.models?.length ? "" : "error"}" data-agent-model-status role="status">${e(desktopModels.modelNotice)}</p>` : ""}
        ${modelRequired ? '<p class="agent-model-status" data-agent-run-status role="status" hidden></p>' : ""}
        <div class="actions">${button(plan ? "编辑大纲" : "编写大纲", "manual-plan", { icon: "edit" })}
          <button class="button primary" type="submit" ${blocked || (modelRequired && !modelSelected) ? "disabled" : ""}>AI 起草大纲</button></div>
        <p class="hint">${desktopRequired ? (desktopAvailable ? "AI 将在桌面客户端中执行，写作要求与所选资料的文本将用于生成大纲。" : "") :
          state.engine?.engine === "demo" ? "演示起草使用本地模板。" : "使用 AI 起草时，写作要求与所选资料的文本将发送至公司模型网关。"}</p>
      </form>
    </section>
    ${plan ? `<section id="plan" class="panel pad stack"><div class="between"><h2 class="section-title">预览并调整大纲</h2><span class="hint">${plan.sections.length} 节</span></div>
      ${outputWarnings(session.outputTelemetry)}
      <p class="hint">每行是一条正文。PowerPoint 自动添加封面。</p><form class="stack" data-form="render">
       <div class="form-grid">
        <label class="field">标题<input required maxlength="110" name="plan_title" value="${e(plan.title)}"></label>
        <label class="field">副标题<input maxlength="140" name="plan_subtitle" value="${e(plan.subtitle)}"></label>
       </div>
       <div class="stack" id="plan-sections">${plan.sections.map((entry, i) => planSection(entry, i)).join("")}</div>
       <div class="actions">${button("添加一节", "add-section", { icon: "plus" })}<button class="button primary" type="submit">${icon("docx")}生成并保存文件</button></div>
      </form></section>` : ""}
    </div><div class="stack sticky-col"><section class="panel pad stack"><h2 class="section-title">文稿文件</h2>
      ${session.result ? fileMarkup(session.result, { edit: true }) : '<p class="hint">尚未保存文件</p>'}
      ${task && session.result ? `<div class="actions">${button("返回任务提交成果", "open-task", { id: task.id, icon: "arrow", cls: "primary" })}</div>` : ""}
    </section></div></div>`;
}

export function planSection(entry, index) {
  const num = String(index + 1).padStart(2, "0");
  return `<div class="plan-card" data-section><div class="between"><span class="plan-card-index">第 ${num} 节</span>
     <button type="button" class="text-button" data-act="remove-section">移除</button></div>
    <input class="form-control" aria-label="段落标题 ${index + 1}" maxlength="80" name="heading" required value="${e(entry.heading)}">
    <textarea class="form-control" aria-label="正文每行一条 ${index + 1}" name="points" required maxlength="4800">${e(entry.points.join("\n"))}</textarea></div>`;
}

export function filterFiles(items, kind = state.filesFilter, query = state.filesQuery) {
  const q = String(query || "").trim().toLocaleLowerCase();
  return (items || []).filter((file) => {
    const ext = (file.extension || "").toLowerCase();
    const matchKind = kind === "all"
      ? true
      : kind === "text"
        ? ["txt", "md"].includes(ext)
        : kind === "image"
          ? ["png", "jpg", "jpeg", "webp", "gif"].includes(ext)
          : ext === kind;
    const matchQuery = !q || file.name.toLocaleLowerCase().includes(q);
    return matchKind && matchQuery;
  });
}

export function renderFilesBulkBar(items) {
  const filtered = filterFiles(items);
  const filteredIds = new Set(filtered.map((f) => f.id));
  const selected = (state.filesSelected || []).filter((id) => filteredIds.has(id));
  const allPicked = filtered.length > 0 && selected.length === filtered.length;
  if (!filtered.length) return "";
  return `<div class="files-bulk-bar">
    <label class="files-select-all">
      <input type="checkbox" data-files-select-all ${allPicked ? "checked" : ""}>
      <span>${allPicked ? "取消全选" : `全选当前列表（${filtered.length}）`}</span>
    </label>
    ${selected.length ? `<div class="files-bulk-actions">
      <span class="hint">已选 <strong>${selected.length}</strong> 项</span>
      <button type="button" class="text-button" data-act="files-clear-select">清空勾选</button>
      <button type="button" class="button sm danger" data-act="files-batch-delete">${icon("trash")}批量删除（${selected.length}）</button>
    </div>` : `<span class="hint">共 ${filtered.length} 个文件 · 可勾选批量删除或单行删除</span>`}
  </div>`;
}

export function renderFilesList(items) {
  const filtered = filterFiles(items);
  if (!filtered.length) {
    return `<p class="empty">${items.length ? "没有匹配的文件" : "当前还没有办公文件"}</p>`;
  }
  const picked = new Set(state.filesSelected || []);
  return `<div class="file-list">${filtered.map((f) => fileMarkup(f, {
    edit: true,
    deletable: true,
    check: "file_select",
    checked: picked.has(f.id)
  })).join("")}</div>`;
}

export function filesView(items) {
  const kinds = [
    ["all", "全部"],
    ["docx", "Word"],
    ["pptx", "PowerPoint"],
    ["image", "图片"],
    ["pdf", "PDF"],
    ["xlsx", "Excel"],
    ["text", "文本"]
  ];
  return `${pageHead("我的文件", "",
    `<div class="actions">${uploadMarkup("上传文件")}${button("新建文稿", "route", { id: "agent", icon: "plus", cls: "primary" })}</div>`)}
  <section class="stack">
    <div class="files-toolbar">
      <div class="tabs" role="group" aria-label="按文件类型筛选">
        ${kinds.map(([key, label]) => `<button type="button" class="tab ${state.filesFilter === key ? "active" : ""}" data-act="files-filter" data-id="${key}">${e(label)}</button>`).join("")}
      </div>
      <div class="files-search">
        <input class="form-control" type="search" data-files-search value="${e(state.filesQuery)}" placeholder="搜索文件名…" aria-label="搜索文件名">
      </div>
    </div>
    <div data-files-bulk>${renderFilesBulkBar(items)}</div>
    <div data-files-list>${renderFilesList(items)}</div>
  </section>`;
}

export function detectDesktopVersion() {
  const direct = typeof window.officeDesktop?.version === "string" ? window.officeDesktop.version.trim() : "";
  if (direct) return direct;
  const ua = typeof navigator !== "undefined" ? String(navigator.userAgent || "") : "";
  const match = ua.match(/(?:qiongyu-office-desktop|QiongyuOffice|穹驭办公)\/([0-9]+\.[0-9]+\.[0-9A-Za-z._-]*)/i);
  return match ? match[1] : "";
}

function formatVersion(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  return /^v/i.test(raw) ? raw : `v${raw}`;
}

export function renderUpdatePanelBody() {
  const isDesktopApp = typeof window.officeDesktop?.runAgent === "function";
  const nativeUpdater = typeof window.officeDesktop?.checkUpdate === "function";
  const info = state.update;

  if (!isDesktopApp) {
    if (info?.full?.url) {
      return `<div class="update-box">
        <div class="update-box-head">
          <strong>桌面客户端 ${e(formatVersion(info.latestVersion))}</strong>
          <span class="update-pill">${e(formatBytes(info.full.size))}</span>
        </div>
        ${info.notes ? `<p class="update-notes">${e(info.notes)}</p>` : ""}
        <div class="actions">
          <a class="button primary" href="${e(info.full.url)}" download>${icon("download")}下载桌面客户端</a>
        </div>
      </div>`;
    }
    return "";
  }

  if (!nativeUpdater) {
    if (info?.hasUpdate && info.full?.url) {
      return `<div class="update-box is-available">
        <div class="update-box-head">
          <strong>发现新版本 ${e(formatVersion(info.latestVersion))}</strong>
          <span class="update-pill">支持增量更新</span>
        </div>
        <p class="update-hint">当前桌面客户端版本较早，升级到 ${e(formatVersion(info.latestVersion))} 后即可支持应用内秒级增量更新与在线下载。</p>
        ${info.notes ? `<p class="update-notes">${e(info.notes)}</p>` : ""}
        <div class="actions">
          <a class="button primary" href="${e(info.full.url)}" download>${icon("download")}下载最新安装包（${e(formatBytes(info.full.size))}）</a>
          ${button("重新检查", "update-check", { icon: "refresh" })}
        </div>
      </div>`;
    }
    return `<div class="update-box">
      <div class="update-box-head">
        <strong>客户端更新</strong>
        <span class="update-meta">${info?.status === "checking" ? "正在检查…" : info?.status === "up-to-date" ? "已是最新" : "可手动检查"}</span>
      </div>
      <div class="actions">${button("检查更新", "update-check", { icon: "refresh", disabled: info?.status === "checking" })}</div>
    </div>`;
  }

  const status = info?.status || "idle";
  const latest = formatVersion(info?.latestVersion);
  const isDelta = info?.mode === "delta";
  const savedPct = isDelta && info?.fullSize > 0 && info?.size > 0 && info.size < info.fullSize
    ? Math.max(1, Math.round((1 - info.size / info.fullSize) * 1000) / 10)
    : 0;
  const modeTag = isDelta
    ? `增量补丁 · ${formatBytes(info.size)}${savedPct ? ` · 省 ${savedPct}% 流量` : ""}`
    : info?.size
      ? `完整安装包 · ${formatBytes(info.size)}`
      : "";

  if (status === "checking") {
    return `<div class="update-box">
      <div class="update-box-head">
        <strong>客户端更新</strong>
        <span class="update-meta">正在检查新版本…</span>
      </div>
      <div class="actions">${button("正在检查…", "update-check", { icon: "refresh", disabled: true })}</div>
    </div>`;
  }

  if (status === "downloading" || status === "verifying") {
    const pct = Math.max(0, Math.min(100, Number(info?.percent) || 0));
    const speed = info?.bytesPerSecond > 0 ? `${formatBytes(info.bytesPerSecond)}/s` : "";
    const progressMeta = status === "verifying"
      ? "正在校验 SHA-256 与补丁完整性…"
      : `${formatBytes(info?.received || 0)} / ${formatBytes(info?.total || info?.size || 0)}${speed ? ` · ${speed}` : ""}`;
    return `<div class="update-box is-active">
      <div class="update-box-head">
        <strong>${status === "verifying" ? `正在校验 ${e(latest)}` : `正在下载 ${e(latest)}`}</strong>
        <span class="update-pill">${e(isDelta ? "增量更新" : "完整更新")} · ${pct}%</span>
      </div>
      <progress class="update-progress" max="100" value="${Math.round(pct)}" aria-label="更新下载进度"></progress>
      <p class="update-meta">${e(progressMeta)}</p>
      ${status === "downloading" ? `<div class="actions">${button("取消下载", "update-cancel", { icon: "x" })}</div>` : ""}
    </div>`;
  }

  if (status === "ready") {
    return `<div class="update-box is-ready">
      <div class="update-box-head">
        <strong>${e(latest)} 已就绪</strong>
        <span class="update-pill is-accent">${e(isDelta ? "增量补丁已校验" : "安装包已校验")}</span>
      </div>
      ${info?.notes ? `<p class="update-notes">${e(info.notes)}</p>` : ""}
      <p class="update-hint">点击后将自动替换更新并重启应用，所有会话与已保存文件不受影响。</p>
      <div class="actions">
        ${button("立即安装并重启", "update-install", { cls: "primary", icon: "refresh" })}
      </div>
    </div>`;
  }

  if (status === "installing") {
    return `<div class="update-box is-active">
      <div class="update-box-head">
        <strong>正在应用 ${e(latest)} 并重启…</strong>
      </div>
      <p class="update-hint">请稍候，穹驭办公将在数秒内自动重新启动。</p>
    </div>`;
  }

  if (status === "available") {
    return `<div class="update-box is-available">
      <div class="update-box-head">
        <strong>发现新版本 ${e(latest)}</strong>
        ${modeTag ? `<span class="update-pill is-accent">${e(modeTag)}</span>` : ""}
      </div>
      ${info?.notes ? `<p class="update-notes">${e(info.notes)}</p>` : ""}
      <div class="actions">
        ${button(`立即下载更新${info?.size ? `（${formatBytes(info.size)}）` : ""}`, "update-download", { cls: "primary", icon: "download" })}
        ${isDelta && info?.hasFull ? button(`改下完整包（${formatBytes(info.fullSize)}）`, "update-download-full", { cls: "ghost" }) : ""}
      </div>
    </div>`;
  }

  if (status === "error") {
    return `<div class="update-box is-error">
      <div class="update-box-head">
        <strong>更新未完成</strong>
      </div>
      <p class="error">${e(info?.error || "检查或下载更新失败，请稍后重试")}</p>
      <div class="actions">
        ${button("重试检查", "update-check", { icon: "refresh" })}
        ${info?.hasFull && info?.latestVersion ? button(`改用完整安装包（${formatBytes(info.fullSize)}）`, "update-download-full", { icon: "download" }) : ""}
        <button type="button" class="button sm" data-act="feedback-quick-upload" data-source="update" data-error="${e(info?.error || "客户端更新失败")}">${icon("upload")}<span>一键上传日志</span></button>
      </div>
    </div>`;
  }

  return `<div class="update-box">
    <div class="update-box-head">
      <strong>客户端更新</strong>
      <span class="update-meta">${status === "up-to-date" ? `已是最新版本（${e(latest || formatVersion(detectDesktopVersion()))}）` : "支持应用内增量更新"}</span>
    </div>
    <div class="actions">${button("检查更新", "update-check", { icon: "refresh" })}</div>
  </div>`;
}

export function settingsView() {
  const isDesktopApp = typeof window.officeDesktop?.runAgent === "function";
  const engine = state.engine?.engine;
  const engineLabel = engine === "demo"
    ? "本地演示模式"
    : engine === "desktop"
      ? (isDesktopApp ? "本机直连 · 独立凭证" : "需在桌面客户端中使用")
      : engine === "dsh"
        ? "企业统一网关 · DeepSeek Harness"
        : "企业统一网关";
  const desktopVersion = isDesktopApp ? formatVersion(detectDesktopVersion()) : "";
  const serverVersion = formatVersion(state.config?.version);
  const webSearch = state.config?.web_search;
  const searchTimeDate = webSearch?.server_time?.date ? ` · 基准 ${webSearch.server_time.date}` : "";
  const webSearchLabel = webSearch?.configured
    ? (webSearch.required ? `DeepSeek Search · 已启用（强制搜索${searchTimeDate}）` : `DeepSeek Search · 已启用${searchTimeDate}`)
    : "DeepSeek Search · 待补入密钥";
  const updateBody = renderUpdatePanelBody();
  return `${pageHead("设置")}
   <div class="settings-grid">
     <section class="panel pad stack">
       <div class="profile-banner">
         <span class="avatar lg">${e(state.me.name.slice(-2))}</span>
         <div>
           <h3>${e(state.me.name)}</h3>
           <p>${state.me.department ? `${e(state.me.department)} · ` : ""}${e(roles[state.me.role])}</p>
         </div>
       </div>
       <div>
         <div class="engine-card-row"><span>所属部门</span><strong>${e(state.me.department || "未设置部门")}</strong></div>
         <div class="engine-card-row"><span>权限角色</span><strong>${e(roles[state.me.role])}</strong></div>
         <div class="engine-card-row"><span>组织时区</span><strong class="mono">${e(state.me.timezone)}</strong></div>
         <div class="engine-card-row"><span>外观主题</span><strong>${getTheme() === "dark" ? "深色 · 白线" : "浅色 · 黑线"}</strong></div>
       </div>
       <div class="actions">
         ${button(getTheme() === "dark" ? "切换浅色模式" : "切换深色模式", "toggle-theme", { icon: getTheme() === "dark" ? "sun" : "moon" })}
         ${button("退出登录", "logout", { icon: "logout" })}
         ${state.config.demo ? button("切换演示身份", "profile") : ""}
         ${typeof window.officeDesktop?.openConnect === "function" ? button("切换工作台地址", "desktop-connect", { icon: "link" }) : ""}
       </div>
     </section>
     <section class="panel pad stack">
       <h2 class="section-title">AI 服务与客户端</h2>
       <div>
         <div class="engine-card-row"><span>运行方式</span><strong>${e(engineLabel)}</strong></div>
         ${state.engine?.model ? `<div class="engine-card-row"><span>默认模型</span><strong class="mono">${e(state.engine.model)}</strong></div>` : ""}
         <div class="engine-card-row"><span>当前环境</span><strong>${isDesktopApp ? "穹驭办公桌面端" : "网页端"}</strong></div>
         <div class="engine-card-row"><span>联网搜索</span><strong>${e(webSearchLabel)}</strong></div>
         ${isDesktopApp ? `<div class="engine-card-row"><span>客户端版本</span><strong class="mono">${e(desktopVersion || "桌面版")}</strong></div>` : ""}
         ${serverVersion ? `<div class="engine-card-row"><span>工作台版本</span><strong class="mono">${e(serverVersion)}</strong></div>` : ""}
       </div>
       <div class="actions">
         ${button("检测服务状态", "check-engine", { icon: "refresh" })}
         ${button("问题反馈 / 上传日志", "feedback-open", { icon: "bug" })}
       </div>
       <div data-update-panel>${updateBody}</div>
     </section>
   </div>`;
}
