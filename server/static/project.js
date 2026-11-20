import { api, button, e, icon, modal, post, toastMessage } from "./ui.js";
import { pageHead } from "./views.js";

const supported = new Set(["docx", "pptx", "pdf", "txt", "md"]);

export async function projectsView(key) {
  if (!key) {
    const projects = await api("/projects");
    return `${pageHead("项目", "",
      `<div class="actions">${button("加入项目", "project-join")}${button("新建项目", "project-create", { cls: "primary", icon: "plus" })}</div>`)}
      ${projects.length
        ? `<div class="project-grid">${projects.map((project) => `<article class="project-card">
            <h2>${e(project.name)}</h2>
            <div class="project-card-foot">
              ${button("开新会话", "route", { id: "chat/project/" + project.id, icon: "message", cls: "sm ghost" })}
              ${button("进入项目", "route", { id: "projects/" + project.id, icon: "right", cls: "sm" })}
            </div>
          </article>`).join("")}</div>`
        : '<p class="empty">暂无项目，可以新建或凭邀请码加入。</p>'}`;
  }

  const [project, files] = await Promise.all([api("/projects/" + encodeURIComponent(key)), api("/files")]);
  const personal = files.filter((file) => !file.assignment_id && supported.has(file.extension));
  return `${pageHead(project.name, "",
    `<div class="actions">${button("返回项目列表", "route", { id: "projects", icon: "left" })}
    ${button("在项目中开新会话", "route", { id: "chat/project/" + project.id, icon: "message", cls: "primary" })}</div>`)}
  <div class="columns"><div class="stack">
    <section class="panel pad stack">
      <h2 class="section-title">发布共享资料</h2>
      <form class="stack" data-form="project-context" data-id="${e(project.id)}">
        <label class="field">资料标题<input name="title" maxlength="110" required></label>
        <label class="field">纪要、要求或决策<textarea name="content" maxlength="5000" rows="3"></textarea></label>
        <label class="field">从我的文件共享<select name="artifact_id"><option value="">不附加文件</option>
          ${personal.map((file) => `<option value="${e(file.id)}">${e(file.name)}</option>`).join("")}</select></label>
        <div class="actions"><button class="button primary" type="submit">发布到项目</button></div>
      </form>
    </section>
    <section class="panel pad stack">
      <h2 class="section-title">共享资料</h2>
      ${project.context.map((item) => `<article class="project-context-entry">
        <div class="between"><h3>${e(item.title)}</h3><span class="hint">${e(item.author)}</span></div>
        ${item.content ? `<p class="prose">${e(item.content)}</p>` : ""}
        ${item.artifact_id ? `<a class="text-button" href="/api/v1/files/${encodeURIComponent(item.artifact_id)}/download" download>${icon("download")}${e(item.filename)}</a>` : ""}
      </article>`).join("") || '<p class="hint">暂无共享资料</p>'}
    </section>
  </div><div class="stack sticky-col"><section class="panel pad stack">
    <h2 class="section-title">项目成员</h2>
    <div class="member-chips">${project.members.map((member) => `<span class="member-chip">${e(member.name)}</span>`).join("")}</div>
    ${project.invite_code ? `<label class="field">邀请码
      <div class="invite-row">
        <input value="${e(project.invite_code)}" readonly aria-label="项目邀请码">
        ${button("复制", "copy-invite", { id: project.invite_code, icon: "copy", cls: "sm" })}
      </div></label>
      <div class="actions">${button("更新邀请码", "project-rotate", { id: project.id, cls: "sm" })}</div>` : ""}
  </section></div></div>`;
}

export function projectDialog(kind) {
  const create = kind === "create";
  modal(create ? "新建项目" : "加入项目", `<form class="stack" data-form="project-${kind}">
    <label class="field">${create ? "项目名称" : "邀请码"}
      <input name="${create ? "name" : "code"}" maxlength="${create ? 80 : 24}" required>
    </label><div class="actions"><button class="button primary" type="submit">${create ? "创建" : "加入"}</button></div>
  </form>`);
}

export async function submitProject(type, id, values) {
  if (type === "project-context") {
    await post("/projects/" + encodeURIComponent(id) + "/context", {
      title: values.title, content: values.content, artifact_id: values.artifact_id || null,
    });
    return id;
  }
  const create = type === "project-create";
  const project = await post(create ? "/projects" : "/projects/join", create ? { name: values.name } : { code: values.code.trim() });
  return project.id;
}

export async function updateInvite(key, render) {
  await post("/projects/" + encodeURIComponent(key) + "/invite", {});
  await render(true);
  toastMessage("邀请码已更新", "success");
}
