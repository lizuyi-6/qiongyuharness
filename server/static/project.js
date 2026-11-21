import { api, button, cleanErrorMessage, deadline, dialog, e, fileIcon, formatBytes, go, icon, labels, materializeFile, modal, post, state, toastMessage, uploadOneFile, uuid } from "./ui.js";
import { pageHead } from "./views.js";
import { renderMarkdown } from "./markdown.js";
import { recordClientLog } from "./feedback.js";

export const PROJECT_DOC_EXTS = new Set(["docx", "pptx", "xlsx", "pdf", "txt", "md", "csv"]);

export const CATEGORY_META = {
  goal: { label: "背景与目标", tag: "目标", tone: "goal" },
  decision: { label: "纪要与决策", tag: "决策", tone: "decision" },
  rule: { label: "口径与规范", tag: "规范", tone: "rule" },
  doc: { label: "方案与资料", tag: "资料", tone: "doc" },
  data: { label: "数据与报表", tag: "数据", tone: "data" }
};

// 按 projectId 追踪正在直传并预整理的文件队列
const projectPendingUploads = new Map();
let activeCategoryFilter = "all";
let currentProjectCache = null;

export function getCachedProject(projectId) {
  if (currentProjectCache && (!projectId || currentProjectCache.id === projectId)) {
    return currentProjectCache;
  }
  return null;
}

function pendingListForProject(projectId) {
  if (!projectId) return [];
  return projectPendingUploads.get(projectId) || [];
}

function categoryBadge(category) {
  const meta = CATEGORY_META[category] || CATEGORY_META.doc;
  return `<span class="project-cat-badge is-${e(meta.tone)}">${e(meta.label)}</span>`;
}

function formatCharCount(chars) {
  const n = Number(chars) || 0;
  if (n <= 0) return "0 字";
  if (n >= 10000) return `约 ${(n / 10000).toFixed(1)} 万字`;
  return `约 ${n.toLocaleString("zh-CN")} 字`;
}

function renderPendingUploadsHtml(projectId) {
  const list = pendingListForProject(projectId);
  if (!list.length) return "";
  return `<div class="project-pending-list" role="status" aria-label="正在上传并预整理的项目文件">
    ${list.map((item) => {
      const isErr = item.status === "error";
      const sizeText = item.size ? formatBytes(item.size) : "";
      return `<div class="project-pending-item ${isErr ? "is-error" : "is-uploading"}">
        <span class="project-pending-main">
          ${isErr ? icon("alert") : '<span class="attached-spinner" aria-hidden="true"></span>'}
          <strong>${e(item.name)}</strong>
          ${sizeText ? `<small class="hint">${e(sizeText)}</small>` : ""}
          <span class="project-pending-status ${isErr ? "is-error" : ""}">${e(isErr ? (item.error || "上传失败") : "正在上传并自动预整理大纲与要点…")}</span>
        </span>
        <button type="button" class="icon-button sm" data-act="project-cancel-upload" data-project-id="${e(projectId)}" data-id="${e(item.tempId)}" title="${isErr ? "关闭提示" : "取消上传"}">${icon("x")}</button>
      </div>`;
    }).join("")}
  </div>`;
}

function syncProjectPendingDom(projectId) {
  for (const slot of document.querySelectorAll(`[data-project-pending="${projectId}"]`)) {
    slot.innerHTML = renderPendingUploadsHtml(projectId);
  }
}

export function cancelProjectPendingUpload(projectId, tempId) {
  const list = pendingListForProject(projectId);
  const target = list.find((it) => it.tempId === tempId);
  if (target?.controller) {
    try { target.controller.abort(); } catch {}
  }
  const next = list.filter((it) => it.tempId !== tempId);
  if (next.length) projectPendingUploads.set(projectId, next);
  else projectPendingUploads.delete(projectId);
  syncProjectPendingDom(projectId);
}

export async function uploadFilesToProject(projectId, rawFiles, { category = "auto", pinned = false, onComplete = null } = {}) {
  if (!projectId) throw new Error("未指定目标项目");
  const files = Array.from(rawFiles || []);
  if (!files.length) return [];
  if (files.length > 12) throw new Error("单次最多同时上传 12 份项目资料");

  const entries = [];
  for (const file of files) {
    const name = String(file?.name || "文件");
    const ext = name.includes(".") ? name.split(".").pop().toLowerCase() : "";
    if (!PROJECT_DOC_EXTS.has(ext)) {
      throw new Error(`「${name}」格式不支持，项目知识库支持 Word、PPT、Excel、PDF、TXT、Markdown 和 CSV`);
    }
    entries.push({
      file,
      tempId: uuid(),
      name,
      extension: ext,
      size: typeof file.size === "number" ? file.size : 0,
      status: "uploading",
      error: "",
      controller: new AbortController()
    });
  }

  projectPendingUploads.set(projectId, [...pendingListForProject(projectId), ...entries]);
  syncProjectPendingDom(projectId);

  const removeEntry = (tempId) => {
    const cur = pendingListForProject(projectId).filter((it) => it.tempId !== tempId);
    if (cur.length) projectPendingUploads.set(projectId, cur);
    else projectPendingUploads.delete(projectId);
  };
  const markError = (tempId, msg) => {
    projectPendingUploads.set(
      projectId,
      pendingListForProject(projectId).map((it) => (it.tempId === tempId ? { ...it, status: "error", error: msg, controller: null } : it))
    );
  };

  // 第一时间将所有拖入文件固化到内存，防止 Electron 拖放句柄失效
  const readyList = [];
  for (const entry of entries) {
    if (entry.controller.signal.aborted) continue;
    try {
      const ready = await materializeFile(entry.file);
      readyList.push({ entry, ready });
    } catch (err) {
      const msg = cleanErrorMessage(err, `读取「${entry.name}」失败`);
      markError(entry.tempId, msg);
      syncProjectPendingDom(projectId);
      recordClientLog("error", "project.upload", msg, { name: entry.name });
      toastMessage(msg, "error");
    }
  }

  const createdItems = [];
  for (const { entry, ready } of readyList) {
    if (entry.controller.signal.aborted) {
      removeEntry(entry.tempId);
      syncProjectPendingDom(projectId);
      continue;
    }
    try {
      const created = await uploadOneFile(ready, {
        target: `/projects/${encodeURIComponent(projectId)}/files`,
        extraFields: {
          category: category || "auto",
          pinned: pinned ? "1" : "0"
        },
        signal: entry.controller.signal,
        preMaterialized: true
      });
      removeEntry(entry.tempId);
      syncProjectPendingDom(projectId);
      createdItems.push(created);
    } catch (err) {
      if (entry.controller.signal.aborted || err?.name === "AbortError") {
        removeEntry(entry.tempId);
        syncProjectPendingDom(projectId);
        continue;
      }
      const msg = cleanErrorMessage(err, `上传「${entry.name}」失败`);
      markError(entry.tempId, msg);
      syncProjectPendingDom(projectId);
      recordClientLog("error", "project.upload", msg, { name: entry.name });
      toastMessage(msg, "error");
    }
  }

  if (createdItems.length) {
    toastMessage(
      createdItems.length > 1
        ? `已上传并自动预整理 ${createdItems.length} 份资料，已纳入项目 AI 上下文`
        : `「${createdItems[0].title}」已上传并完成预整理，已纳入项目 AI 上下文`,
      "success"
    );
    if (typeof onComplete === "function") {
      await onComplete(createdItems);
    }
  }
  return createdItems;
}

function renderContextItemCard(project, item) {
  const canDelete = project.is_owner || item.author_id === state.me?.id;
  const digest = item.digest || {};
  const stats = digest.stats || {};
  const outline = Array.isArray(digest.outline) ? digest.outline : [];
  const keyPoints = Array.isArray(digest.key_points) ? digest.key_points : [];
  const enabled = item.enabled !== false;
  const pinned = Boolean(item.pinned);
  const sizeText = item.size ? formatBytes(item.size) : "";
  const unitLabel = stats.unit_label || (item.extension ? (labels[item.extension] || item.extension) : "上下文纪要");
  const charsLabel = stats.chars ? formatCharCount(stats.chars) : "";

  const showKeyPoints = keyPoints.filter((kp) => !item.content || !item.content.includes(kp)).slice(0, 4);

  return `<article class="project-context-card ${enabled ? "is-enabled" : "is-disabled"} ${pinned ? "is-pinned" : ""}" data-context-item="${e(item.id)}">
    <div class="project-context-card-top">
      <div class="project-context-badges">
        ${categoryBadge(item.category)}
        ${pinned ? `<span class="project-pin-badge" title="置顶资料将优先注入 AI 上下文">${icon("pin")}置顶优先</span>` : ""}
        <h3 class="project-context-title">${e(item.title)}</h3>
      </div>
      <div class="project-context-status-actions">
        <button type="button" class="project-ai-toggle ${enabled ? "is-on" : "is-off"}" data-act="project-context-toggle" data-project-id="${e(project.id)}" data-id="${e(item.id)}" data-enabled="${enabled ? "1" : "0"}" title="${enabled ? "当前已纳入项目 AI 上下文（点击可暂停注入）" : "当前已暂停注入 AI 上下文（点击重新启用）"}">
          ${enabled ? `${icon("checkPlain")}<span>已纳入 AI 上下文</span>` : `<span>已暂停引用 · 点击启用</span>`}
        </button>
      </div>
    </div>

    ${item.content ? `<div class="project-context-note"><p class="prose">${e(item.content)}</p></div>` : ""}

    ${(outline.length || showKeyPoints.length) ? `<div class="project-digest-box">
      <div class="project-digest-head">
        <span class="project-digest-tag">${icon("sparkles")}<span>自动预整理摘要</span></span>
        <span class="hint">${e([unitLabel, charsLabel].filter(Boolean).join(" · "))}</span>
      </div>
      ${outline.length ? `<div class="project-outline-row">
        <small class="project-digest-sublabel">目录结构：</small>
        <div class="project-outline-pills">${outline.map((h) => `<span class="project-outline-pill">${e(h)}</span>`).join("")}</div>
      </div>` : ""}
      ${showKeyPoints.length ? `<ul class="project-keypoints-list">
        ${showKeyPoints.map((pt) => `<li>${e(pt)}</li>`).join("")}
      </ul>` : ""}
    </div>` : ""}

    ${item.artifact_id ? `<div class="project-file-bar">
      <div class="project-file-bar-main">
        ${fileIcon(item.extension)}
        <div class="project-file-bar-meta">
          <strong>${e(item.filename || item.title)}</strong>
          <small>${e([labels[item.extension] || item.extension, unitLabel, charsLabel, sizeText].filter((v, i, a) => v && a.indexOf(v) === i).join(" · "))}</small>
        </div>
      </div>
      <div class="project-file-bar-actions">
        <button type="button" class="text-button" data-act="preview" data-id="${e(item.artifact_id)}" data-name="${e(item.filename || item.title)}">${icon("search")}预览提取正文</button>
        <a class="text-button" href="/api/v1/files/${encodeURIComponent(item.artifact_id)}/download" download>${icon("download")}下载原文件</a>
      </div>
    </div>` : ""}

    <div class="project-context-card-foot">
      <span class="hint">由 ${e(item.author || "项目成员")} 添加 · ${e(deadline(item.updated_at || item.created_at))}</span>
      <div class="project-context-tools">
        <button type="button" class="text-button" data-act="project-context-pin" data-project-id="${e(project.id)}" data-id="${e(item.id)}" data-pinned="${pinned ? "1" : "0"}">${icon("pin")}${pinned ? "取消置顶" : "置顶优先"}</button>
        <button type="button" class="text-button" data-act="project-context-edit" data-project-id="${e(project.id)}" data-id="${e(item.id)}">${icon("edit")}编辑</button>
        ${canDelete ? `<button type="button" class="text-button is-danger" data-act="project-context-delete" data-project-id="${e(project.id)}" data-id="${e(item.id)}" data-title="${e(item.title)}">${icon("trash")}移除</button>` : ""}
      </div>
    </div>
  </article>`;
}

export function renderProjectContextList(project, filter = activeCategoryFilter) {
  const items = Array.isArray(project?.context) ? project.context : [];
  if (!items.length) {
    return `<div class="project-empty-context">
      ${icon("layers")}
      <strong>项目知识库暂无上下文资料</strong>
      <p class="hint">直接把电脑上的 Word、PPT、Excel、PDF、Markdown 或 CSV 拖入本页，系统会自动提取章节大纲与核心要点并注入 AI 会话。</p>
    </div>`;
  }
  const filtered = items.filter((it) => {
    if (filter === "all") return true;
    if (filter === "enabled") return it.enabled !== false;
    return it.category === filter;
  });
  if (!filtered.length) {
    return `<p class="hint">当前分类下暂无项目资料，可切换到「全部」或直接拖入新文件。</p>`;
  }
  return `<div class="project-context-stack">${filtered.map((item) => renderContextItemCard(project, item)).join("")}</div>`;
}

export function filterProjectContext(filter) {
  activeCategoryFilter = filter || "all";
  if (!currentProjectCache) return;
  const listSlot = document.querySelector("[data-project-context-list]");
  if (listSlot) listSlot.innerHTML = renderProjectContextList(currentProjectCache, activeCategoryFilter);
  for (const tab of document.querySelectorAll("[data-act=project-filter]")) {
    const on = tab.dataset.id === activeCategoryFilter;
    tab.classList.toggle("active", on);
    tab.setAttribute("aria-pressed", String(on));
  }
}

export async function projectsView(key) {
  if (!key) {
    const projects = await api("/projects");
    return `${pageHead(
      "项目空间",
      "把项目背景、会议决策、统一口径与多格式参考文档归档在一起，自动预整理大纲并注入项目下的每一次 AI 对话。",
      `<div class="actions">${button("凭邀请码加入", "project-join")}${button("新建项目", "project-create", { cls: "primary", icon: "plus" })}</div>`
    )}
    ${projects.length
      ? `<div class="project-grid">${projects.map((project) => {
          const enabledCount = Number(project.enabled_context_count || 0);
          const totalCount = Number(project.context_count || 0);
          const fileCount = Number(project.file_count || 0);
          const memberCount = Number(project.member_count || 1);
          return `<article class="project-card">
            <div class="project-card-head">
              <div class="project-card-title-row">
                <h2>${e(project.name)}</h2>
                <span class="project-role-tag">${project.is_owner ? "创建者" : "成员"}</span>
              </div>
              <p class="project-card-desc">${e(project.description || "暂未填写项目背景说明，进入项目可上传资料并自动预整理总览。")}</p>
            </div>
            <div class="project-card-stats">
              <span class="project-stat-pill ${enabledCount > 0 ? "is-active" : ""}">
                ${icon(enabledCount > 0 ? "checkPlain" : "layers")}
                <span>${enabledCount > 0 ? `已启用 ${enabledCount} 项 AI 上下文` : "待添加上下文"}</span>
              </span>
              <span class="project-stat-meta">${fileCount} 份文档 · ${totalCount} 项资料 · ${memberCount} 位成员</span>
            </div>
            <div class="project-card-foot">
              ${button("带项目上下文开会话", "route", { id: "chat/project/" + project.id, icon: "message", cls: "sm ghost" })}
              ${button("管理资料与预整理", "route", { id: "projects/" + project.id, icon: "right", cls: "sm primary" })}
            </div>
          </article>`;
        }).join("")}</div>`
      : `<div class="panel pad"><div class="project-empty-context">
          ${icon("layers")}
          <strong>暂无项目空间</strong>
          <p class="hint">新建项目后，可直接拖入多份文档（Word / PPT / Excel / PDF / MD / CSV），系统将自动预整理章节与核心结论，并在每次项目对话中自动提供给 AI。</p>
          <div class="actions">${button("凭邀请码加入", "project-join")}${button("新建第一个项目", "project-create", { cls: "primary", icon: "plus" })}</div>
        </div></div>`}`;
  }

  const [project, files] = await Promise.all([
    api("/projects/" + encodeURIComponent(key)),
    api("/files")
  ]);
  currentProjectCache = project;
  const personal = files.filter((file) => !file.assignment_id && PROJECT_DOC_EXTS.has(file.extension));
  const overview = project.overview || {};
  const stats = overview.stats || {
    total_count: project.context.length,
    enabled_count: project.context.filter((c) => c.enabled !== false).length,
    pinned_count: project.context.filter((c) => c.pinned && c.enabled !== false).length,
    file_count: project.context.filter((c) => c.artifact_id).length,
    note_count: project.context.filter((c) => !c.artifact_id).length,
    total_chars: 0,
    category_counts: {}
  };
  const catCounts = stats.category_counts || {};
  if (activeCategoryFilter !== "all" && activeCategoryFilter !== "enabled" && !catCounts[activeCategoryFilter]) {
    activeCategoryFilter = "all";
  }
  const effectiveBrief = (overview.effective_brief || "").trim();
  const hasDesktopAi = typeof window.officeDesktop?.generateTitle === "function" || typeof window.officeDesktop?.runAgent === "function";

  const filterTabs = [
    ["all", "全部资料", stats.total_count],
    ["enabled", "✓ 已注入 AI", stats.enabled_count],
    ["goal", "背景与目标", catCounts.goal || 0],
    ["decision", "纪要与决策", catCounts.decision || 0],
    ["rule", "口径与规范", catCounts.rule || 0],
    ["doc", "方案与资料", catCounts.doc || 0],
    ["data", "数据与报表", catCounts.data || 0]
  ];

  return `<div class="project-workspace" data-project-dropzone="${e(project.id)}">
    <div class="project-drop-overlay" aria-hidden="true">
      <div class="drop-overlay-box">
        ${icon("upload")}
        <strong>松开鼠标，直接上传到「${e(project.name)}」并自动预整理</strong>
        <span>自动提取正文结构、识别分类标签并加入项目 AI 上下文（支持 Word、PPT、Excel、PDF、MD、TXT、CSV）</span>
      </div>
    </div>

    ${pageHead(
      project.name,
      project.description || "把项目文档、会议纪要与规范口径归档在此，自动预整理后对项目下所有会话生效。",
      `<div class="actions">
        ${button("返回项目列表", "route", { id: "projects", icon: "left" })}
        ${button("编辑项目目标", "project-edit-info", { id: project.id, icon: "edit" })}
        ${button("在项目中开新会话", "route", { id: "chat/project/" + project.id, icon: "message", cls: "primary" })}
      </div>`
    )}

    <section class="project-context-meter panel pad">
      <div class="project-meter-main">
        <div class="project-meter-badge ${stats.enabled_count > 0 ? "is-ready" : "is-empty"}">
          ${icon(stats.enabled_count > 0 ? "checkPlain" : "sparkles")}
          <strong>${stats.enabled_count > 0 ? `AI 上下文已就绪 · 已启用 ${stats.enabled_count} / ${stats.total_count} 项` : "尚未添加项目上下文"}</strong>
        </div>
        <p class="project-meter-copy">
          ${stats.enabled_count > 0
            ? `包含 <strong>${stats.file_count}</strong> 份文档与 <strong>${stats.note_count}</strong> 条纪要规范（置顶优先 <strong>${stats.pinned_count}</strong> 项，预整理共 <strong>${e(formatCharCount(stats.total_chars))}</strong>）。在本项目开启的任何会话中，AI 都会<strong>自动读取预整理总览与已启用资料</strong>，无需每轮重复上传。`
            : "拖入项目文件或填写会议决策后，系统会自动预整理目录结构与核心要点。凡是在本项目中发起的会话，AI 均会自动读取这些上下文。"}
        </p>
      </div>
      <div class="project-meter-actions">
        ${button("从我的文件批量选用", "project-pick-files", { id: project.id, icon: "folder", cls: "sm" })}
        ${button("在项目中开会话", "route", { id: "chat/project/" + project.id, icon: "message", cls: "sm primary" })}
      </div>
    </section>

    <div class="columns">
      <div class="stack">
        <section class="panel pad stack" aria-label="上传文件与添加项目上下文">
          <div class="between">
            <div>
              <h2 class="section-title">添加项目资料与上下文</h2>
              <p class="hint">方式一：直接拖拽或选择本地文件批量上传（自动命名、自动分类并预整理）；方式二：录入会议纪要、目标或口径要求。</p>
            </div>
            ${button("从我的文件批量导入", "project-pick-files", { id: project.id, icon: "folder", cls: "sm" })}
          </div>

          <label class="project-upload-dropzone" data-project-upload-box>
            <input type="file" data-project-file-input data-project-id="${e(project.id)}" accept=".docx,.pptx,.xlsx,.pdf,.txt,.md,.csv" multiple>
            <div class="project-upload-icon">${icon("upload")}</div>
            <div class="project-upload-copy">
              <strong>点击选择本地文件，或直接将文件拖拽到此页面任意位置</strong>
              <span>支持多选批量直传 · Word (.docx)、PPT (.pptx)、Excel (.xlsx)、PDF (.pdf)、Markdown (.md)、TXT、CSV · 入库即自动预整理大纲与要点</span>
            </div>
            <span class="button sm primary project-upload-cta">${icon("plus")}选择本地文件</span>
          </label>

          <div data-project-pending="${e(project.id)}">${renderPendingUploadsHtml(project.id)}</div>

          <details class="project-manual-details" ${!project.context.length ? "open" : ""}>
            <summary class="project-manual-summary">
              <span>${icon("edit")}<strong>录入项目背景、会议决策、统一口径或关联已有文件</strong></span>
              <small class="hint">点击展开 / 收起表单</small>
            </summary>
            <form class="stack project-context-form" data-form="project-context" data-id="${e(project.id)}">
              <div class="form-grid">
                <label class="field">上下文分类
                  <select name="category" aria-label="上下文分类">
                    <option value="auto">✨ 自动识别分类（推荐）</option>
                    <option value="goal">🎯 背景与目标（项目立项、目标、受众）</option>
                    <option value="decision">📌 纪要与决策（会议结论、排期、负责人）</option>
                    <option value="rule">📏 口径与规范（数据口径、术语、写作要求）</option>
                    <option value="doc">📄 方案与资料（参考文稿、汇报材料）</option>
                    <option value="data">📊 数据与报表（业务数据、统计清单）</option>
                  </select>
                </label>
                <label class="field">资料标题（选填，选文件后可自动生成）
                  <input name="title" maxlength="110" placeholder="如：Q4 试点评审结论 / 统一数据口径">
                </label>
              </div>
              <label class="field">纪要、决策、口径说明或补充背景
                <textarea name="content" maxlength="8000" rows="3" placeholder="填写希望 AI 在本项目所有会话中遵守的决策、背景、数据口径或资料摘录…"></textarea>
              </label>
              <div class="form-grid">
                <label class="field">或从「我的文件」附加单份文档（选填）
                  <select name="artifact_id" aria-label="从我的文件共享">
                    <option value="">不附加文件（仅作为文本上下文）</option>
                    ${personal.map((file) => `<option value="${e(file.id)}">${e(file.name)}（${e(labels[file.extension] || file.extension)}）</option>`).join("")}
                  </select>
                </label>
                <label class="field project-pin-check-wrap">
                  <span>注入优先级</span>
                  <label class="project-pin-check">
                    <input type="checkbox" name="pinned" value="1">
                    <span>置顶优先（始终排在 AI 上下文最前）</span>
                  </label>
                </label>
              </div>
              <div class="actions">
                <button class="button primary" type="submit">${icon("sparkles")}发布并自动预整理</button>
              </div>
            </form>
          </details>
        </section>

        <section class="panel pad stack project-overview-panel" aria-label="项目预整理总览">
          <div class="between">
            <div class="project-overview-title">
              <span class="project-digest-tag">${icon("sparkles")}<span>项目知识库预整理总览</span></span>
              <h2 class="section-title">AI 上下文结构化大纲</h2>
            </div>
            <div class="actions">
              ${button("重新预整理", "project-preorganize", { id: project.id, icon: "refresh", cls: "sm" })}
              ${hasDesktopAi && project.context.length ? button("AI 深度提炼", "project-ai-preorganize", { id: project.id, icon: "sparkles", cls: "sm" }) : ""}
              ${button("编辑总览", "project-edit-brief", { id: project.id, icon: "edit", cls: "sm ghost" })}
            </div>
          </div>
          <p class="hint">系统已将本项目的目标说明、置顶决策与各份文档的目录要点预整理为以下结构化总览，AI 每次回答或写文档前都会优先读取：</p>
          ${effectiveBrief
            ? `<div class="project-brief-box md-body">${renderMarkdown(effectiveBrief)}</div>`
            : `<div class="project-brief-empty">
                <p class="hint">当前暂无预整理总览。上传文件或添加纪要后将自动生成，也可点击右上角「重新预整理」或「编辑总览」手动编写项目总纲。</p>
              </div>`}
        </section>

        <section class="panel pad stack" aria-label="共享资料与上下文列表">
          <div class="between">
            <h2 class="section-title">项目上下文资料库（${stats.total_count}）</h2>
            <span class="hint">可随时切换「已纳入 AI 上下文」开关或置顶重要决策</span>
          </div>
          <div class="tabs project-filter-tabs" role="tablist" aria-label="按分类筛选项目上下文">
            ${filterTabs.map(([id, label, count]) => `
              <button type="button" class="tab ${activeCategoryFilter === id ? "active" : ""}" data-act="project-filter" data-id="${e(id)}" aria-pressed="${activeCategoryFilter === id}">
                ${e(label)}<span class="tab-count">${count}</span>
              </button>`).join("")}
          </div>
          <div data-project-context-list>${renderProjectContextList(project, activeCategoryFilter)}</div>
        </section>
      </div>

      <div class="stack sticky-col">
        <section class="panel pad stack">
          <div class="between">
            <h2 class="section-title">项目专属会话</h2>
            ${button("新会话", "route", { id: "chat/project/" + project.id, icon: "plus", cls: "sm primary" })}
          </div>
          <p class="hint">从此处开启或进入的会话，均自动挂载本项目已启用的 ${stats.enabled_count} 项预整理上下文。</p>
          ${project.conversations?.length
            ? `<div class="project-conv-list">${project.conversations.map((conv) => `
                <button type="button" class="project-conv-item" data-act="route" data-id="chat/${e(conv.id)}">
                  ${icon(conv.pinned ? "pin" : "message")}
                  <span class="project-conv-meta">
                    <strong>${e(conv.title)}</strong>
                    <small>${e(deadline(conv.updated_at))}</small>
                  </span>
                  ${icon("right")}
                </button>`).join("")}</div>`
            : `<p class="hint">你在本项目中暂无历史会话，点击右上角「新会话」即可基于项目资料开始创作。</p>`}
        </section>

        <section class="panel pad stack">
          <h2 class="section-title">项目成员（${project.members.length}）</h2>
          <div class="member-chips">${project.members.map((member) =>
            `<span class="member-chip" title="${e([member.department, rolesLabel(member.role)].filter(Boolean).join(" · "))}">${e(member.name)}</span>`
          ).join("")}</div>
          ${project.invite_code ? `<label class="field">成员协作邀请码
            <div class="invite-row">
              <input value="${e(project.invite_code)}" readonly aria-label="项目邀请码">
              ${button("复制", "copy-invite", { id: project.invite_code, icon: "copy", cls: "sm" })}
            </div></label>
            <div class="actions">${button("更换邀请码", "project-rotate", { id: project.id, cls: "sm" })}</div>` : ""}
        </section>
      </div>
    </div>
  </div>`;
}

function rolesLabel(role) {
  const map = { admin: "超级管理员", manager: "部门主管", employee: "员工", reviewer: "审阅人" };
  return map[role] || "";
}

export function projectDialog(kind) {
  const create = kind === "create";
  modal(create ? "新建项目空间" : "凭邀请码加入项目", `<form class="stack" data-form="project-${kind}">
    <label class="field">${create ? "项目名称" : "24 位项目邀请码"}
      <input name="${create ? "name" : "code"}" maxlength="${create ? 80 : 24}" required placeholder="${create ? "如：2026 Q2 渠道增长专项" : "粘贴项目创建者分享的邀请码"}">
    </label>
    ${create ? `<label class="field">项目目标与背景说明（选填，将自动纳入项目 AI 上下文总览）
      <textarea name="description" rows="3" maxlength="2000" placeholder="简述该项目的核心目标、交付标准、目标受众或统一口径要求…"></textarea>
    </label>` : ""}
    <div class="actions"><button class="button primary" type="submit">${create ? "创建项目空间" : "加入项目"}</button></div>
  </form>`);
}

export function openEditProjectInfoModal(projectId) {
  const project = getCachedProject(projectId) || state.chatWorkspace?.project;
  if (!project) return;
  modal("编辑项目目标与背景", `<form class="stack" data-form="project-edit-info" data-id="${e(project.id)}">
    <label class="field">项目名称
      <input name="name" maxlength="80" required value="${e(project.name)}" ${project.is_owner ? "" : "readonly"}>
    </label>
    <label class="field">项目目标与背景说明（AI 在本项目所有会话中均会优先读取）
      <textarea name="description" rows="5" maxlength="2000" placeholder="填写项目的核心背景、阶段目标、关键业务指标或统一写作口径…">${e(project.description || "")}</textarea>
    </label>
    <div class="actions"><button class="button primary" type="submit">保存项目信息</button></div>
  </form>`);
}

export function openEditProjectBriefModal(projectId) {
  const project = getCachedProject(projectId) || state.chatWorkspace?.project;
  if (!project) return;
  const overview = project.overview || {};
  const currentText = (project.brief || overview.effective_brief || "").trim();
  modal("编辑项目预整理总览", `<form class="stack" data-form="project-edit-brief" data-id="${e(project.id)}">
    <p class="hint">可在此调整或完善项目预整理总览（支持 Markdown）。保存后，本项目下的所有 AI 会话将直接以该总览作为项目全局上下文。</p>
    <label class="field">项目预整理总览正文
      <textarea name="brief" rows="12" maxlength="12000" placeholder="### 一、项目背景与核心目标\n...\n### 二、核心决策与执行口径\n...">${e(currentText)}</textarea>
    </label>
    <div class="actions">
      <button class="button primary" type="submit">保存总览</button>
    </div>
  </form>`);
}

export function openEditContextItemModal(projectId, itemId) {
  const project = getCachedProject(projectId) || state.chatWorkspace?.project;
  const item = (project?.context || []).find((c) => c.id === itemId);
  if (!project || !item) return;
  modal("编辑项目上下文资料", `<form class="stack" data-form="project-context-edit" data-project-id="${e(project.id)}" data-id="${e(item.id)}">
    <div class="form-grid">
      <label class="field">上下文分类
        <select name="category" aria-label="上下文分类">
          ${Object.entries(CATEGORY_META).map(([k, meta]) =>
            `<option value="${e(k)}" ${item.category === k ? "selected" : ""}>${e(meta.label)}</option>`
          ).join("")}
        </select>
      </label>
      <label class="field">资料标题
        <input name="title" maxlength="110" required value="${e(item.title)}">
      </label>
    </div>
    <label class="field">纪要、决策、口径说明或资料解读${item.artifact_id ? "（选填，留空则使用文件自动预整理摘要）" : ""}
      <textarea name="content" rows="5" maxlength="8000" ${item.artifact_id ? "" : "required"} placeholder="补充该资料的核心结论、适用范围或执行要求…">${e(item.content || "")}</textarea>
    </label>
    <div class="form-grid">
      <label class="project-pin-check">
        <input type="checkbox" name="pinned" value="1" ${item.pinned ? "checked" : ""}>
        <span>置顶优先（在 AI 上下文中排在最前）</span>
      </label>
      <label class="project-pin-check">
        <input type="checkbox" name="enabled" value="1" ${item.enabled !== false ? "checked" : ""}>
        <span>纳入项目 AI 上下文</span>
      </label>
    </div>
    <div class="actions"><button class="button primary" type="submit">保存并更新预整理</button></div>
  </form>`);
}

export async function openBatchPickFilesModal(projectId) {
  const [project, files] = await Promise.all([
    getCachedProject(projectId) ? Promise.resolve(getCachedProject(projectId)) : api("/projects/" + encodeURIComponent(projectId)),
    api("/files")
  ]);
  const existingArtifactIds = new Set((project?.context || []).map((c) => c.artifact_id).filter(Boolean));
  const personal = files.filter((file) => !file.assignment_id && PROJECT_DOC_EXTS.has(file.extension));
  modal(`从「我的文件」批量添加到项目`, `<form class="stack" data-form="project-batch-files" data-id="${e(projectId)}">
    <p class="hint">勾选要共享到「${e(project?.name || "当前项目")}」的办公文档，系统将自动提取各文档章节结构与要点并纳入项目 AI 上下文：</p>
    <div class="form-grid">
      <label class="field">默认分类
        <select name="category" aria-label="默认分类">
          <option value="auto">✨ 按文件名与内容自动识别分类</option>
          <option value="goal">🎯 背景与目标</option>
          <option value="decision">📌 纪要与决策</option>
          <option value="rule">📏 口径与规范</option>
          <option value="doc">📄 方案与资料</option>
          <option value="data">📊 数据与报表</option>
        </select>
      </label>
      <label class="field project-pin-check-wrap">
        <span>优先级</span>
        <label class="project-pin-check">
          <input type="checkbox" name="pinned" value="1">
          <span>批量设为置顶优先</span>
        </label>
      </label>
    </div>
    <div class="file-list project-batch-file-list">
      ${personal.length
        ? personal.map((file) => {
            const already = existingArtifactIds.has(file.id);
            const size = file.size ? formatBytes(file.size) : "";
            return `<label class="file project-pick-file-row ${already ? "is-already" : ""}">
              <input type="checkbox" name="artifact_ids" value="${e(file.id)}" ${already ? "checked disabled" : ""}>
              ${fileIcon(file.extension)}
              <span class="file-meta">
                <strong>${e(file.name)}</strong>
                <small>${e(labels[file.extension] || file.extension)}${size ? " · " + size : ""}</small>
              </span>
              ${already ? `<span class="project-cat-badge is-doc">已在项目中</span>` : ""}
            </label>`;
          }).join("")
        : `<p class="hint">「我的文件」中暂无可共享的独立文档。你可以直接将电脑上的文件拖拽到项目页面上传。</p>`}
    </div>
    <div class="actions">
      <button class="button primary" type="submit" ${personal.some((f) => !existingArtifactIds.has(f.id)) ? "" : "disabled"}>${icon("sparkles")}导入所选文件并自动预整理</button>
    </div>
  </form>`);
}

export async function openQuickProjectContextModal(projectId) {
  if (!projectId) return;
  const [project, files] = await Promise.all([
    api("/projects/" + encodeURIComponent(projectId)),
    api("/files").catch(() => [])
  ]);
  currentProjectCache = project;
  const personal = (files || []).filter((file) => !file.assignment_id && PROJECT_DOC_EXTS.has(file.extension));
  const stats = project.overview?.stats || {};

  modal(`管理项目上下文 · ${project.name}`, `<div class="stack" data-quick-project-modal="${e(project.id)}">
    <div class="project-quick-banner">
      <span class="project-digest-tag">${icon("checkPlain")}<span>当前已启用 ${stats.enabled_count ?? project.context.length} 项上下文（自动注入本会话）</span></span>
      <button type="button" class="text-button" data-act="route" data-id="projects/${e(project.id)}">进入完整项目空间${icon("right")}</button>
    </div>

    <label class="project-upload-dropzone is-compact" data-project-upload-box>
      <input type="file" data-project-file-input data-project-id="${e(project.id)}" accept=".docx,.pptx,.xlsx,.pdf,.txt,.md,.csv" multiple>
      <div class="project-upload-icon">${icon("upload")}</div>
      <div class="project-upload-copy">
        <strong>直接上传本地文件到该项目（自动预整理大纲与要点）</strong>
        <span>支持 Word、PPT、Excel、PDF、MD、TXT、CSV 多文件直传，上传后立即对当前会话生效</span>
      </div>
      <span class="button sm primary project-upload-cta">${icon("plus")}选择文件</span>
    </label>
    <div data-project-pending="${e(project.id)}">${renderPendingUploadsHtml(project.id)}</div>

    <form class="stack project-context-form" data-form="project-context" data-id="${e(project.id)}">
      <div class="form-grid">
        <label class="field">分类
          <select name="category" aria-label="上下文分类">
            <option value="auto">✨ 自动识别分类</option>
            <option value="goal">🎯 背景与目标</option>
            <option value="decision">📌 纪要与决策</option>
            <option value="rule">📏 口径与规范</option>
            <option value="doc">📄 方案与资料</option>
            <option value="data">📊 数据与报表</option>
          </select>
        </label>
        <label class="field">标题（选文件后可自动生成）
          <input name="title" maxlength="110" placeholder="如：核心业务口径 / 评审会议纪要">
        </label>
      </div>
      <label class="field">上下文正文（纪要、决策、目标或口径要求）
        <textarea name="content" rows="3" maxlength="8000" placeholder="输入希望 AI 在该项目所有对话中自动遵循的背景、结论或要求…"></textarea>
      </label>
      <div class="form-grid">
        <label class="field">或关联「我的文件」中的文档（选填）
          <select name="artifact_id" aria-label="从我的文件共享">
            <option value="">不附加文件</option>
            ${personal.map((file) => `<option value="${e(file.id)}">${e(file.name)}</option>`).join("")}
          </select>
        </label>
        <label class="field project-pin-check-wrap">
          <span>优先级</span>
          <label class="project-pin-check">
            <input type="checkbox" name="pinned" value="1" checked>
            <span>置顶优先注入 AI</span>
          </label>
        </label>
      </div>
      <div class="actions">
        <button class="button primary" type="submit">${icon("sparkles")}添加到项目上下文并预整理</button>
      </div>
    </form>
  </div>`);
}

export async function openBindProjectModal(currentProjectId = "") {
  const projects = await api("/projects");
  modal("关联或切换项目上下文", `<form class="stack" data-form="chat-bind-project">
    <p class="hint">为当前会话关联项目后，该项目下已预整理的背景目标、会议决策、规范口径与共享文档将自动注入 AI 上下文，无需逐一上传文件。</p>
    ${projects.length
      ? `<div class="project-bind-list">
          <label class="project-bind-option ${!currentProjectId ? "active" : ""}">
            <input type="radio" name="project_id" value="" ${!currentProjectId ? "checked" : ""}>
            <div class="project-bind-copy">
              <strong>不关联项目（独立个人会话）</strong>
              <small>仅使用当前会话中直接上传或勾选的参考资料</small>
            </div>
          </label>
          ${projects.map((p) => {
            const on = p.id === currentProjectId;
            return `<label class="project-bind-option ${on ? "active" : ""}">
              <input type="radio" name="project_id" value="${e(p.id)}" ${on ? "checked" : ""}>
              <div class="project-bind-copy">
                <div class="between">
                  <strong>${e(p.name)}</strong>
                  <span class="project-stat-pill ${p.enabled_context_count > 0 ? "is-active" : ""}">${p.enabled_context_count || 0} 项生效上下文 · ${p.file_count || 0} 份文档</span>
                </div>
                <small>${e(p.description || "自动载入该项目的预整理总览与全部已启用共享资料")}</small>
              </div>
            </label>`;
          }).join("")}
        </div>
        <div class="actions">
          ${button("新建项目", "project-create", { icon: "plus" })}
          <button class="button primary" type="submit">应用所选项目上下文</button>
        </div>`
      : `<div class="project-empty-context">
          ${icon("layers")}
          <strong>你还没有创建或加入任何项目</strong>
          <p class="hint">创建项目后，可以把多份参考文档与要求归档预整理，并在多个会话间复用。</p>
          <div class="actions">
            ${button("凭邀请码加入", "project-join")}
            ${button("立即新建项目", "project-create", { cls: "primary", icon: "plus" })}
          </div>
        </div>`}
  </form>`);
}

export async function submitProject(type, id, values, form = null) {
  if (type === "project-context") {
    await post("/projects/" + encodeURIComponent(id) + "/context", {
      title: values.title || "",
      content: values.content || "",
      artifact_id: values.artifact_id || null,
      category: values.category || "auto",
      pinned: Boolean(values.pinned)
    });
    return id;
  }
  if (type === "project-edit-info") {
    const updated = await api("/projects/" + encodeURIComponent(id), {
      method: "PATCH",
      body: { name: values.name, description: values.description || "" }
    });
    currentProjectCache = updated;
    return id;
  }
  if (type === "project-edit-brief") {
    const updated = await api("/projects/" + encodeURIComponent(id), {
      method: "PATCH",
      body: { brief: values.brief || "" }
    });
    currentProjectCache = updated;
    return id;
  }
  if (type === "project-context-edit") {
    const projectId = form?.dataset?.projectId || state.route[1] || state.chatWorkspace?.project?.id;
    await api(`/projects/${encodeURIComponent(projectId)}/context/${encodeURIComponent(id)}`, {
      method: "PATCH",
      body: {
        title: values.title,
        content: values.content || "",
        category: values.category || "doc",
        pinned: Boolean(values.pinned),
        enabled: Boolean(values.enabled)
      }
    });
    return projectId;
  }
  if (type === "project-batch-files") {
    const data = form ? new FormData(form) : null;
    const ids = data ? data.getAll("artifact_ids") : (values.artifact_ids ? [values.artifact_ids] : []);
    if (!ids.length) throw new Error("请至少勾选一份要导入的文档");
    await post(`/projects/${encodeURIComponent(id)}/context/batch-files`, {
      artifact_ids: ids,
      category: values.category || "auto",
      pinned: Boolean(values.pinned)
    });
    return id;
  }
  const create = type === "project-create";
  const project = await post(
    create ? "/projects" : "/projects/join",
    create ? { name: values.name, description: values.description || "" } : { code: values.code.trim() }
  );
  return project.id;
}

export async function toggleProjectContextItem(projectId, itemId, nextEnabled, rerender) {
  await api(`/projects/${encodeURIComponent(projectId)}/context/${encodeURIComponent(itemId)}`, {
    method: "PATCH",
    body: { enabled: Boolean(nextEnabled) }
  });
  toastMessage(nextEnabled ? "已纳入项目 AI 上下文" : "已暂停该资料注入 AI 上下文", "info");
  if (typeof rerender === "function") await rerender(projectId);
}

export async function togglePinProjectContextItem(projectId, itemId, nextPinned, rerender) {
  await api(`/projects/${encodeURIComponent(projectId)}/context/${encodeURIComponent(itemId)}`, {
    method: "PATCH",
    body: { pinned: Boolean(nextPinned), ...(nextPinned ? { enabled: true } : {}) }
  });
  toastMessage(nextPinned ? "已设为置顶优先上下文" : "已取消置顶", "info");
  if (typeof rerender === "function") await rerender(projectId);
}

export async function deleteProjectContextItem(projectId, itemId, title, rerender) {
  await api(`/projects/${encodeURIComponent(projectId)}/context/${encodeURIComponent(itemId)}`, {
    method: "DELETE"
  });
  toastMessage(`已从项目移除「${title || "该资料"}」`, "success");
  if (typeof rerender === "function") await rerender(projectId);
}

export async function runProjectPreorganize(projectId, { useAi = false, triggerButton = null, rerender = null } = {}) {
  if (!projectId) return;
  const origHtml = triggerButton?.innerHTML || "";
  if (triggerButton?.isConnected) {
    triggerButton.disabled = true;
    triggerButton.innerHTML = `${icon("refresh")}<span>正在预整理…</span>`;
  }
  try {
    const res = await post(`/projects/${encodeURIComponent(projectId)}/preorganize`, {});
    currentProjectCache = res;
    if (useAi && res?.ai_prepare && typeof window.officeDesktop?.runAgent === "function") {
      if (triggerButton?.isConnected) {
        triggerButton.innerHTML = `${icon("sparkles")}<span>AI 正在深度提炼总览…</span>`;
      }
      try {
        const modelsInfo = await window.officeDesktop.models?.().catch(() => null);
        const pickModel = modelsInfo?.models?.[0] || "";
        if (pickModel) {
          const aiRes = await window.officeDesktop.runAgent({
            system: res.ai_prepare.system,
            prompt: res.ai_prepare.prompt,
            model: pickModel
          });
          const aiText = String(aiRes?.text || "").trim();
          if (aiText && aiText.length >= 20) {
            const patched = await api(`/projects/${encodeURIComponent(projectId)}`, {
              method: "PATCH",
              body: { brief: aiText.slice(0, 12000) }
            });
            currentProjectCache = patched;
          }
        }
      } catch {
        // AI 提炼失败时保留服务端结构化预整理总览，不阻断用户流程
      }
    }
    toastMessage(useAi ? "项目知识库已完成 AI 深度预整理" : "项目资料与结构化总览已重新预整理", "success");
    if (typeof rerender === "function") await rerender(projectId);
  } finally {
    if (triggerButton?.isConnected) {
      triggerButton.disabled = false;
      triggerButton.innerHTML = origHtml;
    }
  }
}

export async function updateInvite(key, render) {
  await post("/projects/" + encodeURIComponent(key) + "/invite", {});
  await render(true);
  toastMessage("邀请码已更新", "success");
}
