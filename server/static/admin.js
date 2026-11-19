import { api, badge, button, deadline, e, eventLabels, fileMarkup, fromLocal,
  icon, isManager, localInput, modal, post, priorities, roles, state, uploadSelected, uuid } from "./ui.js";
import { pageHead, pagination } from "./views.js";

const fields = {scheduled:"待发布",pending:"待开始",in_progress:"进行中",paused:"已暂停",submitted:"待审核",
  returned:"退回修改",approved:"已验收",overdue:"已逾期",revoked:"已撤销",closed:"已关闭"};
const pageSize = 20;
let roster = [];
let plans = [];
let createProgress = null;

export function clearAdmin() {
  roster = [];
  plans = [];
  createProgress = null;
}

function options(values, chosen) {
  return Object.entries(values).map(([k,v])=>`<option value="${e(k)}" ${k===chosen?"selected":""}>${e(v)}</option>`).join("");
}

export async function taskAdmin(reviewOnly = false) {
  const filters = state.admin;
  const p = new URLSearchParams({offset:filters.offset,limit:pageSize});
  if (reviewOnly) p.set("status","submitted");
  else if (filters.status) p.set("status",filters.status);
  if (filters.department) p.set("department",filters.department);
  if (filters.date) p.set("work_date",filters.date);
  const [items, users, fetchedPlans] = await Promise.all([
    api("/admin/tasks?"+p),
    isManager()?api("/people"):[],
    !reviewOnly && isManager()?api("/admin/plans"):[]
  ]);
  roster = users;
  plans = fetchedPlans;
  const departments = Object.fromEntries([...new Set(users.map((v)=>v.department).filter(Boolean))].map((v)=>[v,v]));
  const planList = fetchedPlans.map((row)=>`<article><div class="between"><strong>${e(row.data.title)}</strong>${badge(row.state)}</div>
     <p class="hint">${e(row.data.work_date)}${row.data.end_date?" 至 "+e(row.data.end_date):""}
      · ${row.data.assignee_ids.length} 人${row.data.repeat==="weekdays"?" · 按工作日":""}</p>
     <div class="actions">${row.state==="draft"||row.state==="active"?button("编辑", "edit-plan",{id:row.id,icon:"edit"}):""}
      ${row.state==="draft"?button("发布", "publish-plan",{id:row.id,cls:"primary"}):""}</div></article>`).join("");
  return `${pageHead(reviewOnly?"审核队列":"任务管理","",
     !reviewOnly && isManager()?button("新建任务","new-task",{cls:"primary",icon:"plus"}):"")}
   <div class="stack admin-sections"><section aria-label="${reviewOnly?"待审成果":"任务列表"}">
    <form class="filter" data-form="admin-filter" aria-label="筛选任务">
    ${!reviewOnly?`<select name="status" aria-label="任务状态"><option value="">全部状态</option>${options(fields,filters.status)}</select>`:""}
    ${Object.keys(departments).length?`<select name="department" aria-label="部门"><option value="">全部部门</option>${options(departments,filters.department)}</select>`:""}
    <input type="date" name="date" aria-label="工作日期" value="${e(filters.date)}">
    <button class="button" type="submit">${icon("search")}筛选</button>${button("清空筛选","admin-reset")}</form>
    <div class="table-wrap"><table class="admin-list"><thead><tr><th>任务名称</th><th>经办人</th><th>部门</th><th>工作日期</th><th>截止时间</th><th>优先级</th><th>状态</th><th>操作</th></tr></thead>
    <tbody>${items.map(t=>`<tr><td><strong>${e(t.title)}</strong></td><td>${e(t.assignee_name)}</td><td>${e(t.department)}</td>
     <td>${e(t.work_date)}</td><td>${e(deadline(t.due_at))}</td><td>${e(priorities[t.priority])}</td>
     <td>${badge(t.status)} ${t.overdue?badge("overdue"):""}</td><td><button type="button" class="text-button" data-act="open-task" data-id="${e(t.id)}">${t.status==="submitted"?"查看成果":"查看任务"}</button></td></tr>`).join("") ||
      `<tr><td colspan="8"><p class="empty">${reviewOnly?"暂无待审核的成果":"当前筛选没有匹配任务"}</p></td></tr>`}</tbody></table></div>
    ${pagination("admin",filters.offset,pageSize,items.length)}
   </section>
   ${!reviewOnly && isManager()?`<section class="plans-section"><h2 class="section-title">草稿与重复计划</h2>
   <div class="plan-list">${planList || '<p class="hint">暂无任务草稿或重复计划。</p>'}</div></section>`:""}</div>`;
}

function reviewers(chosen, source) {
  return '<option value="">按管理权限审核</option>' +
    options(Object.fromEntries(source.filter((p)=>p.role!=="employee").map((p)=>[p.id,p.name+" · "+roles[p.role]])), chosen);
}

export async function showTaskForm(id = "") {
  if (!roster.length) roster = await api("/people");
  createProgress = null;
  const row = plans.find((f)=>f.id===id);
  const data = row?.data || {};
  const picked = new Set(data.assignee_ids || []);
  const isActive = row?.state==="active";
  const work = data.work_date || state.me.today;
  modal(row ? isActive ? "编辑重复计划" : "编辑任务草稿" : "新建任务",
    `<form class="stack" data-form="task-template" data-key="${uuid()}" data-publish-key="${uuid()}" data-id="${e(id)}" data-version="${row?.version || 0}">
    <div class="form-grid">
      <label class="field span">任务标题<input required name="title" maxlength="130" value="${e(data.title)}" placeholder="例如：制作项目方案汇报 PPT"></label>
      <label class="field span">任务要求<textarea name="description" rows="4" maxlength="8000" placeholder="写明背景、具体要求与注意事项">${e(data.description)}</textarea></label>
      <label class="field span">交付标准<textarea name="deliverable" rows="2" maxlength="2000" placeholder="交付格式、模板和验收要求">${e(data.deliverable)}</textarea></label>
      <label class="field">成果类型<select name="kind">${options({docx:"Word",pptx:"PowerPoint",other:"综合"},data.kind || "docx")}</select></label>
      <label class="field">优先级<select name="priority">${options(priorities,data.priority || "normal")}</select></label>
      <label class="field">工作日期<input required type="date" name="work_date" value="${e(work)}" ${isActive?"readonly":""}></label>
      <label class="field">截止时间（组织时区）<input type="datetime-local" name="due_at" value="${e(localInput(data.due_at))}"></label>
      <label class="field">定时发布（选填）<input type="datetime-local" name="publish_at" value="${e(localInput(data.publish_at))}"></label>
      <label class="field">审核人<select name="reviewer_id">${reviewers(data.reviewer_id,roster)}</select></label>
      <label class="field">重复方式<select name="repeat" ${isActive?"disabled":""}>${options({once:"单次",weekdays:"按工作日"},data.repeat || "once")}</select></label>
      <label class="field" data-repeat ${data.repeat==="weekdays"?"":"hidden"}>计划结束日期<input name="end_date" type="date" value="${e(data.end_date)}"></label>
      <label class="field span" data-repeat ${data.repeat==="weekdays"?"":"hidden"}>排除日期（逗号分隔，YYYY-MM-DD）<input name="exclude_dates" value="${e(data.exclude_dates?.join(", "))}"></label>
      <label class="field span" data-repeat ${data.repeat==="weekdays"?"":"hidden"}>额外工作日（逗号分隔，YYYY-MM-DD）<input name="include_dates" value="${e(data.include_dates?.join(", "))}"></label>
      <fieldset class="field span"><legend>分派员工</legend><div class="check-list">
      ${roster.filter(p=>["manager","employee"].includes(p.role)).map(p=>`<label><input name="assignee_ids" type="checkbox" value="${e(p.id)}" ${picked.has(p.id)?"checked":""}>${e(p.name)} · ${e(p.department)}</label>`).join("")}
      </div></fieldset>
      ${isActive?"":`<label class="field span">任务资料（选填）<input name="materials" type="file" multiple accept=".docx,.pptx,.xlsx,.pdf,.txt,.md"></label>`}
    </div>
    ${row?.files?.length?`<div class="file-list">${row.files.map(file=>fileMarkup(file)).join("")}</div>`:""}
    <div class="actions"><button type="submit" class="button" name="intent" value="draft">${row?"保存修改":"保存草稿"}</button>
      ${isActive?"":'<button type="submit" class="button primary" name="intent" value="publish">发布任务</button>'}</div>
  </form>`);
}

function valuesFor(form) {
  const values=new FormData(form);
  const dates=(k)=>String(values.get(k) || "").split(/[,，\s]+/).map(x=>x.trim()).filter(Boolean);
  const payload={title:values.get("title"),description:values.get("description"),deliverable:values.get("deliverable"),
    kind:values.get("kind"),priority:values.get("priority"),work_date:values.get("work_date"),
    due_at:fromLocal(values.get("due_at")),publish_at:fromLocal(values.get("publish_at")),
    reviewer_id:values.get("reviewer_id") || null,assignee_ids:values.getAll("assignee_ids"),
    repeat:values.get("repeat") || "weekdays",end_date:values.get("end_date") || null,exclude_dates:dates("exclude_dates"),include_dates:dates("include_dates")};
  if (!payload.assignee_ids.length) throw new Error("请至少选择一位员工");
  return {payload, files:[...form.querySelector("[name=materials]")?.files || []]};
}

export async function saveTaskForm(form, intent) {
  const {payload, files}=valuesFor(form);
  const serial = JSON.stringify(payload);
  let result;
  if (createProgress && createProgress.serial===serial) result=createProgress.result;
  else {
    if (form.dataset.id) {
      result=await api("/tasks/"+encodeURIComponent(form.dataset.id),{method:"PATCH",body:{...payload,version:Number(form.dataset.version)}});
    } else {
      if (createProgress) form.dataset.key=uuid();
      result=await post("/tasks",payload,form.dataset.key);
    }
    form.dataset.id=result.id;
    form.dataset.version=result.version;
    createProgress={result,serial,finishedFiles:new Set(),publishKey:uuid()};
  }
  for (const file of files) {
    const signature=[file.name,file.size,file.lastModified].join("|");
    if (createProgress.finishedFiles.has(signature)) continue;
    await uploadSelected([file],"","/tasks/"+encodeURIComponent(result.id)+"/materials");
    createProgress.finishedFiles.add(signature);
  }
  if (intent==="publish") {
    await post("/tasks/"+encodeURIComponent(result.id)+"/publish",{version:result.version},createProgress.publishKey);
  }
  createProgress=null;
}

export function showPublish(id) {
  const row=plans.find((x)=>x.id===id);
  if (!row) throw new Error("任务草稿已变更，请刷新后重试");
  modal("发布任务",`<form class="stack" data-form="publish-plan" data-id="${e(id)}" data-version="${row.version}" data-key="${uuid()}">
    <p class="prose">${e(row.data.title)}</p><p class="hint">将按任务设定和员工名单发布。</p>
    <div class="actions"><button type="submit" class="button primary">确认发布</button></div></form>`);
}

export async function changeAssignment(t) {
  if (!roster.length) roster=await api("/people");
  const coworkers=roster.filter((p)=>["employee","manager"].includes(p.role));
  modal("调整任务",`<form data-form="change-task" class="stack">
   <div class="form-grid">
   <label class="field span">任务标题<input name="title" required maxlength="130" value="${e(t.title)}"></label>
   <label class="field span">任务要求<textarea name="description" rows="3" maxlength="8000">${e(t.description)}</textarea></label>
   <label class="field span">交付标准<textarea name="deliverable" rows="2" maxlength="2000">${e(t.deliverable)}</textarea></label>
   <label class="field">截止时间<input type="datetime-local" name="due_at" value="${e(localInput(t.due_at))}"></label>
   <label class="field">优先级<select name="priority">${options(priorities,t.priority)}</select></label>
   ${["pending","scheduled"].includes(t.status)?`<label class="field">经办员工<select name="assignee_id">${options(Object.fromEntries(coworkers.map(p=>[p.id,p.name+" · "+p.department])),t.assignee_id)}</select></label>`:""}
   <label class="field">审核人<select name="reviewer_id">${reviewers(t.reviewer_id,roster)}</select></label>
   </div><div class="actions"><button type="submit" class="button primary">保存调整</button></div></form>`);
}

export function assignmentValues(form, version) {
  const data=Object.fromEntries(new FormData(form));
  return {...data,version,due_at:fromLocal(data.due_at),reviewer_id:data.reviewer_id || null};
}

export async function auditPage() {
  const items=await api("/admin/audit?offset="+state.auditOffset+"&limit="+pageSize);
  return `${pageHead("审计记录","查看任务、附件和模型会话的操作记录。")}
    <section class="panel"><div class="table-wrap"><table>
      <thead><tr><th>时间</th><th>操作者</th><th>事件</th><th>关联任务</th><th>摘要</th></tr></thead>
      <tbody>${items.map(x=>`<tr><td>${e(deadline(x.created_at))}</td><td>${e(x.name || "系统")}</td>
        <td>${e(eventLabels[x.type] || x.type)}</td><td>${e(x.title || "—")}</td>
        <td>${e(x.payload.usage?.total_tokens !== undefined ? x.payload.usage.total_tokens + " tokens" :
          x.payload.revision ? "第 " + x.payload.revision + " 版" : x.payload.sha256 ? "SHA-256 " + x.payload.sha256.slice(0,12) : x.payload.engine || "—")}</td></tr>`).join("") ||
        '<tr><td colspan="5"><p class="empty">暂无记录</p></td></tr>'}</tbody></table></div>
    ${pagination("audit",state.auditOffset,pageSize,items.length)}</section>`;
}
