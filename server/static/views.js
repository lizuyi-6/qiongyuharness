import { active, api, badge, button, dateLabel, deadline, e, eventLabels, fileIcon, fileMarkup,
  icon, isManager, isReview, labels, priorities, readDraft, roles, state, uploadMarkup } from "./ui.js";
import { conversationNav } from "./chat.js";

export const views = {
  today:["今日任务","today"], agent:["文稿","docx"], files:["我的文件","folder"],
  chat:["办公会话","message"],
  projects:["项目","folder"],
  history:["任务历史","history"], admin:["任务管理","tasks"], review:["审核队列","check"],
  audit:["审计记录","audit"], settings:["设置","sliders"], task:["任务详情","docx"]
};
export const demoUsers = [
  ["lin","林知夏","员工 · 产品与运营"], ["admin","李祖祎","超级管理员"],
  ["manager","周婧","主管 · 产品与运营"], ["chen","陈以宁","员工 · 产品与运营"],
  ["song","宋予安","员工 · 技术研发"], ["reviewer","顾言","审阅人 · 产品与运营"]
];

const empty = (copy) => `<p class="empty">${e(copy)}</p>`;
const navItem = (id) => {
  const [route] = state.route;
  const on = route === id;
  const [name, glyph] = views[id];
  return `<button class="nav ${on?"active":""}" type="button" data-act="route" data-id="${id}" title="${name}" ${on?'aria-current="page"':""}>
    ${icon(glyph)}<span>${name}</span></button>`;
};
const brand = () => `<span class="brand-mark">${icon("mark")}</span><span class="brand-name">穹驭办公</span>`;

export function shell(content) {
  const person = state.me;
  const current = views[state.route[0]]?.[0] || "工作台";
  const offline = !navigator.onLine;
  return `<div class="layout ${state.route[0]==="chat"?"workspace-route":""}">
  <aside class="sidebar" aria-label="应用导航">
    <div class="brand">${brand()}</div>
    <nav>${conversationNav(state.chatWorkspace || {conversations:[]})}
      <div class="nav-group"><p class="nav-label">办公</p>${["today","projects","history","files","agent"].map(navItem).join("")}</div>
      ${isReview()?`<div class="nav-group"><p class="nav-label">管理</p>${isManager()?navItem("admin"):""}${navItem("review")}${person.role==="admin"?navItem("audit"):""}</div>`:""}
    </nav>
    <div class="sidebar-foot">${navItem("settings")}
      <button type="button" class="identity" data-act="profile" aria-label="账户与身份">
        <span class="avatar">${e(person.name.slice(-2))}</span><span class="id-copy"><strong>${e(person.name)}</strong><small>${e(person.department)} · ${e(roles[person.role])}</small></span>${icon("down")}
      </button>
    </div>
  </aside>
  <div class="topbar"><div class="crumbs"><span>${e(person.department)}</span>${icon("right")}<strong>${current}</strong></div>
    <div class="top-actions">
      ${state.config.demo?'<button class="demo-mark" type="button" data-act="profile">演示环境</button>':""}
      <button class="search-button" type="button" data-act="search" aria-label="搜索任务">${icon("search")}<span>搜索任务</span><kbd>Ctrl K</kbd></button>
      <button class="icon-button" type="button" data-act="refresh" aria-label="刷新" title="刷新">${icon("refresh")}</button>
      <span class="connection-state ${offline?"offline":""}"><i class="dot"></i>${offline?"离线":"已连接"}</span>
    </div>
  </div><main class="content" id="workplace">${content}</main></div>`;
}

export function loginView(error = "") {
  return `<div class="login"><section class="login-content"><div class="brand">${brand()}</div>
    <h1>${state.config?.demo?"选择演示身份":"登录穹驭办公"}</h1>
    ${error?`<p class="error">${e(error)}</p>`:""}
    ${state.config?.demo?demoUsers.map(([id,name,role])=>button(name+" · "+role,"demo-login",{id,cls:id==="lin"?"primary":""})).join(""):
      state.config?.sso?button("使用公司账号登录","sso",{cls:"primary"}):'<p class="notice warn">企业登录入口尚未配置。</p>'}
  </section></div>`;
}

export function pageHead(title, description = "", action = "") {
  return `<header class="page-head"><div><h1>${e(title)}</h1>${description?`<p>${e(description)}</p>`:""}</div>${action}</header>`;
}

function taskTable(items, emptyMessage) {
  return `<div class="table-wrap"><table class="work-list">
    <thead><tr><th scope="col">任务</th><th scope="col">状态</th><th scope="col">截止时间</th></tr></thead>
    <tbody>${items.length ? items.map((task) => `<tr><td>
      <button class="task-link" type="button" data-act="open-task" data-id="${e(task.id)}" aria-label="${e(task.title)}">
        ${fileIcon(task.kind)}<span><strong>${e(task.title)}</strong><small>${e(labels[task.kind])}${task.priority!=="normal"?` · ${e(priorities[task.priority])}`:""}</small></span>
      </button>
    </td><td>${badge(task.status)}</td><td class="${task.overdue?"late":""}">
      ${e(deadline(task.due_at))}${task.overdue?'<small class="deadline-note">已逾期</small>':""}
    </td></tr>`).join("") : `<tr><td colspan="3">${empty(emptyMessage)}</td></tr>`}</tbody>
  </table></div>`;
}

function taskSubset(items, key) {
  if (key==="working") return items.filter((x)=>["pending","in_progress","paused"].includes(x.status));
  if (key==="done") return items.filter((x)=>["submitted","approved"].includes(x.status));
  if (key==="attention") return items.filter((x)=>x.overdue||x.status==="returned");
  return items;
}

export function dashboard() {
  const {today, filter} = state;
  const ordered = [...today].sort((a,b) => {
    if (Boolean(a.overdue) !== Boolean(b.overdue)) return Number(Boolean(b.overdue)) - Number(Boolean(a.overdue));
    if (Boolean(a.due_at) !== Boolean(b.due_at)) return Number(Boolean(b.due_at)) - Number(Boolean(a.due_at));
    return (a.due_at || "").localeCompare(b.due_at || "");
  });
  const reminders = ordered.filter((task)=>active(task.status) && task.due_at).slice(0,3);
  const tabs = [["all","全部"],["working","待办理"],["done","已提交／验收"],["attention","退回／逾期"]];
  const date = dateLabel(state.me.today, {month:"long",day:"numeric",weekday:"long"});
  return `${pageHead("今日任务","",`<time class="head-date" datetime="${e(state.me.today)}">${icon("calendar")}${e(date)}</time>`)}
  <div class="dashboard">
    <section class="daily-list" aria-label="任务列表">
      <div class="tabs" aria-label="任务状态">${tabs.map(([key,title]) => `<button class="tab ${filter===key?"active":""}" type="button" data-act="filter-today" data-id="${key}" aria-pressed="${filter===key}">${title}<span class="tab-count">${taskSubset(today,key).length}</span></button>`).join("")}</div>
      ${taskTable(taskSubset(ordered,filter), "当前分类暂无任务")}
    </section>
    <aside class="deadlines" aria-labelledby="deadlines-title"><h2 id="deadlines-title">截止提醒</h2>
      ${reminders.length ? `<ul class="agenda">${reminders.map((task)=>`<li>
        <button class="agenda-item" type="button" data-act="open-task" data-id="${e(task.id)}" aria-label="查看截止任务：${e(task.title)}">
          <strong>${e(task.title)}</strong><span class="${task.overdue?"late":""}">${e(deadline(task.due_at))}${task.overdue?" · 已逾期":""}</span>
        </button></li>`).join("")}</ul>` : '<p class="hint">暂无即将到期的任务</p>'}
    </aside>
  </div>`;
}

export function historyView(items, offset, pageSize) {
  return `${pageHead("任务历史")}<section aria-label="已发布任务">
    ${taskTable(items, "暂无历史任务")}
    ${pagination("history",offset,pageSize,items.length)}
  </section>`;
}

export function pagination(prefix, offset, size, found) {
  return `<div class="pagination"><span class="hint">第 ${Math.floor(offset/size)+1} 页</span>
  ${button("上一页",prefix+"-prev",{disabled:offset===0})}${button("下一页",prefix+"-next",{disabled:found<size})}</div>`;
}

function timeline(task) {
  return task.events.map((event) => {
    const changeNames = {title:"标题",description:"要求",deliverable:"交付标准",due_at:"截止时间",priority:"优先级",reviewer_id:"审核人",assignee_id:"员工"};
    const changes = event.type==="task.updated" ? Object.keys(event.payload.after || {}).map((key) => {
      const before = event.payload.before[key];
      const after = event.payload.after[key];
      if (["reviewer_id","assignee_id"].includes(key)) return `${changeNames[key]}已调整`;
      const value = (v)=> key==="due_at" ? deadline(v) : key==="priority" ? priorities[v] : String(v || "—");
      return `${changeNames[key] || key}：${value(before)} → ${value(after)}`;
    }).join("\n") : "";
    return `<div><strong>${e(eventLabels[event.type] || event.type)}</strong>
    <small>${e(event.name || "系统")} · ${e(deadline(event.created_at))}</small>
    ${event.payload.note || changes ? `<p class="prose hint">${e(event.payload.note || changes)}</p>`:""}</div>`;
  }).join("") || "<p class=\"hint\">暂无办理记录</p>";
}

function reviewBox(t, submission, canReview) {
  const showReview = canReview && t.status==="submitted" && submission.status==="submitted";
  return `<div class="review-box"><div class="between"><strong>第 ${submission.revision} 版</strong>${badge(submission.status)}</div>
    <p class="hint">${e(deadline(submission.created_at))}</p>
    ${submission.note?`<p class="prose">${e(submission.note)}</p>`:""}
    ${submission.link?`<a href="${e(submission.link)}" target="_blank" rel="noopener noreferrer">${icon("link")} 打开交付链接</a>`:""}
    ${submission.files.map((f)=>fileMarkup(f)).join("")}
    ${submission.review_note?`<p class="notice ${submission.status==="returned"?"warn":""}">${e(submission.review_note)}</p>`:""}
    ${showReview?`<div class="actions">${button("通过验收","review-task",{cls:"primary",id:submission.id})}
          ${button("退回修改","return-task",{cls:"danger",id:submission.id})}</div>`:""}
  </div>`;
}

export function taskDetail(t) {
  const mine = state.me.id===t.assignee_id;
  const canManage = isManager();
  const canReview = !mine && (canManage || state.me.id===t.reviewer_id);
  const canDo = mine && active(t.status);
  const note = readDraft("submit/"+t.id) || {};
  const picked = new Set(note.artifact_ids || []);
  const late = t.overdue ? badge("overdue") : "";
  return `${pageHead(t.title, "由 " + t.owner_name + " 分派给 " + t.assignee_name,
    `<div class="actions">${button("返回","route",{id:mine?"today":isManager()?"admin":"review",icon:"left"})}
    ${canManage && !["approved","revoked","closed"].includes(t.status)?button("调整任务","change-task",{icon:"edit"}):""}</div>`)}
  <div class="columns"><div class="stack">
    <section class="panel pad stack">
     <div class="summary"><div><span class="hint">任务状态</span>${badge(t.status)} ${late}</div>
     <div><span class="hint">截止时间</span><strong>${e(deadline(t.due_at))}</strong></div>
     <div><span class="hint">优先级</span><strong>${e(priorities[t.priority])}</strong></div></div>
     <section><h2 class="group-title">任务要求</h2><p class="prose">${e(t.description || "暂无补充要求")}</p></section>
     <section><h2 class="group-title">交付标准</h2><p class="prose">${e(t.deliverable || "由经办员工与主管确认")}</p></section>
     ${t.materials.length?`<section><h2 class="group-title">任务资料</h2><div class="file-list">${t.materials.map(f=>fileMarkup(f)).join("")}</div></section>`:""}
     ${canDo?`<div class="actions">${t.status==="in_progress"?button("暂停办理","task-action",{id:"pause"}):button("开始办理","task-action",{id:"start",cls:"primary"})}
         ${button("申请延期","task-note",{id:"extension"})}${button("留言","task-note",{id:"comment"})}</div>`:""}
     ${canManage && !["approved","revoked","closed"].includes(t.status)?`<div class="actions">
         ${button("撤销任务","task-note",{id:"revoke",cls:"danger"})}${t.status!=="scheduled"?button("手工关闭","task-note",{id:"close"}):""}</div>`:""}
    </section>
    ${canDo?`<section class="panel pad stack"><h2 class="section-title">提交成果</h2>
      <div class="actions">${uploadMarkup("上传成果")}
      ${button("从我的文件关联","choose-file",{icon:"folder"})}
      ${button("用会话办理","agent-task",{id:t.id,icon:"message"})}</div>
      <form data-form="submission" class="stack" data-key="${e(note.key || "")}">
        ${t.files.length ? `<div class="file-list">${t.files.map(f=>fileMarkup(f,{check:"artifact_ids",checked:picked.has(f.id),edit:true})).join("")}</div>` : ""}
        <label class="field">成果说明<textarea name="note" rows="3" maxlength="4000" placeholder="简述本次交付内容和需要审核的事项">${e(note.note || "")}</textarea></label>
        <label class="field">交付链接（选填）<input name="link" type="url" maxlength="500" value="${e(note.link || "")}" placeholder="https://"></label>
        <div class="actions"><button class="button primary" type="submit">${icon("check")}提交审核</button></div>
      </form></section>`:""}
    <section class="panel pad stack"><h2 class="section-title">历史交付</h2>${t.submissions.map((sub)=>reviewBox(t,sub,canReview)).join("")||'<p class="hint">暂无提交记录</p>'}</section>
  </div><section class="panel pad stack"><h2 class="section-title">办理记录</h2><div class="timeline">${timeline(t)}</div></section></div>`;
}

export function agentSetup(session, task, files, desktopModels = {}) {
  const plan = session.plan;
  const desktopRequired=state.engine?.client_execution===true;
  const desktopAvailable=typeof window.officeDesktop?.runAgent==="function";
  const modelSelected=Boolean(desktopModels.selectedModel && desktopModels.models?.includes(desktopModels.selectedModel));
  const blocked=desktopRequired && !desktopAvailable;
  const modelRequired=desktopRequired && desktopAvailable;
  const saved = [...(task?.materials || []), ...files.filter((f)=>!f.assignment_id || f.assignment_id===task?.id)]
    .filter((f)=>["docx","pptx","pdf","txt","md"].includes(f.extension)).slice(0,20);
  return `${pageHead("文稿",task?"关联任务 · "+task.title:"",
    task?button("返回任务","open-task",{id:task.id,icon:"left"}):"")}
    <div class="columns"><div class="stack"><section class="panel pad stack">
      <h2 class="section-title">文稿信息</h2>
      <form class="stack" data-form="generate">
        <div class="form-grid">
          <label class="field span">标题<input required name="title" maxlength="110" value="${e(session.title)}" placeholder="例如：产品方案汇报"></label>
          <label class="field">输出格式<select name="kind"><option value="docx" ${session.kind==="docx"?"selected":""}>Word（DOCX）</option><option value="pptx" ${session.kind==="pptx"?"selected":""}>PowerPoint（PPTX）</option></select></label>
          <label class="field" data-pages ${session.kind!=="pptx"?"hidden":""}>页数（含封面）<input name="pages" type="number" min="3" max="20" value="${e(session.pages || 8)}"></label>
          ${modelRequired?`<label class="field span agent-model-picker">模型<select data-agent-model aria-label="执行模型"
            ${desktopModels.models?.length?"":"disabled"}>
            <option value="">${desktopModels.models?.length?"请选择模型":"暂无可用模型"}</option>
            ${(desktopModels.models || []).map(model=>`<option value="${e(model)}" ${desktopModels.selectedModel===model?"selected":""}>${e(model)}</option>`).join("")}
          </select></label>`:""}
          <label class="field span">写作要求<textarea name="instructions" rows="3" maxlength="6000" placeholder="受众、重点和格式要求">${e(session.instructions || "")}</textarea></label>
        </div>
        <div class="draft-sources">
          ${saved.length?`<details ${session.material_ids.length?"open":""}><summary>选择参考资料${session.material_ids.length?`（已选 ${session.material_ids.length} 份）`:""}</summary>
            <div class="file-list">${saved.map(f=>fileMarkup(f,{check:"material_ids",checked:session.material_ids.includes(f.id)})).join("")}</div></details>`:""}
          ${uploadMarkup("上传参考资料")}
        </div>
        ${blocked?'<p class="notice" role="status">请在桌面客户端中使用 AI 办理</p>':""}
        ${modelRequired && desktopModels.modelNotice?`<p class="agent-model-status ${desktopModels.models?.length?"":"error"}" data-agent-model-status role="status">${e(desktopModels.modelNotice)}</p>`:""}
        ${modelRequired?'<p class="agent-model-status" data-agent-run-status role="status" hidden></p>':""}
        <div class="actions">${button(plan?"编辑大纲":"编写大纲","manual-plan",{icon:"edit",cls:"primary"})}
          <button class="button" type="submit" ${blocked || (modelRequired && !modelSelected)?"disabled":""}>AI 起草大纲</button></div>
        <p class="hint">${desktopRequired?(desktopAvailable?"AI 将在桌面客户端中执行，写作要求与所选资料的文本将用于生成大纲。":""):
          state.engine?.engine==="demo"?"演示起草使用本地模板。":"使用 AI 起草时，写作要求与所选资料的文本将发送至公司模型网关。"}</p>
      </form>
    </section>
    ${plan?`<section id="plan" class="panel pad stack"><h2 class="section-title">预览并调整大纲</h2>
      <p class="hint">每行是一条正文。PowerPoint 自动添加封面。</p><form class="stack" data-form="render">
       <div class="form-grid">
        <label class="field">标题<input required maxlength="110" name="plan_title" value="${e(plan.title)}"></label>
        <label class="field">副标题<input maxlength="140" name="plan_subtitle" value="${e(plan.subtitle)}"></label>
       </div>
       <div class="stack" id="plan-sections">${plan.sections.map((entry,i)=>planSection(entry,i)).join("")}</div>
       <div class="actions">${button("添加一节","add-section",{icon:"plus"})}<button class="button primary" type="submit">${icon("docx")}生成并保存文件</button></div>
      </form></section>`:""}
    </div><div class="stack"><section class="panel pad stack"><h2 class="section-title">文稿文件</h2>
      ${session.result?fileMarkup(session.result,{edit:true}):'<p class="hint">尚未保存文件</p>'}
      ${task && session.result?`<div class="actions">${button("返回任务提交成果","open-task",{id:task.id,icon:"arrow"})}</div>`:""}
    </section></div></div>`;
}

export function planSection(entry, index) {
  return `<div class="plan-card" data-section><div class="between"><label class="field">段落标题</label>
     <button type="button" class="text-button" data-act="remove-section">移除</button></div>
    <input class="form-control" aria-label="段落标题 ${index+1}" maxlength="80" name="heading" required value="${e(entry.heading)}">
    <textarea class="form-control" aria-label="正文每行一条 ${index+1}" name="points" required maxlength="4800">${e(entry.points.join("\n"))}</textarea></div>`;
}

export function filesView(items) {
  return `${pageHead("我的文件","",`<div class="actions">${uploadMarkup("上传文件")}${button("新建文稿","route",{id:"agent",icon:"plus",cls:"primary"})}</div>`)}
  <section class="panel pad stack">
    <div class="file-list">${items.map((f)=>fileMarkup(f,{edit:true})).join("")||empty("当前还没有办公文件")}</div>
  </section>`;
}

export function settingsView() {
  return `${pageHead("设置","查看登录身份和当前模型连接。")}
   <div class="columns"><div class="stack"><section class="panel pad stack"><h2 class="section-title">登录账号</h2>
     <div class="summary"><div><span class="hint">姓名</span><strong>${e(state.me.name)}</strong></div><div><span class="hint">部门</span><strong>${e(state.me.department)}</strong></div>
     <div><span class="hint">权限</span><strong>${e(roles[state.me.role])}</strong></div></div>
     <p class="hint">组织时区：${e(state.me.timezone)}</p>
     <div class="actions">${button("退出登录","logout",{icon:"logout"})}${state.config.demo?button("切换演示身份","profile"):""}</div>
   </section><section class="panel pad stack"><h2 class="section-title">模型连接</h2>
     <p class="hint">${state.engine?.engine==="demo"?"本地演示模板":state.engine?.engine==="dsh"?"DeepSeek Harness · 公司模型网关":"公司模型网关"}</p>
     ${state.engine?.model?`<p class="hint">模型：${e(state.engine.model)}</p>`:""}
     ${button("检查连接","check-engine",{icon:"refresh"})}</section></div></div>`;
}
