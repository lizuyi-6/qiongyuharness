import { api, button, e, icon, readDraft, saveDraft, state, toastMessage } from "./ui.js";

export function deckFiles(session) {
  return [...new Map((session?.turns || []).filter((turn) => turn.file?.extension === "pptx")
    .map((turn) => [turn.file.id, turn.file])).values()].reverse();
}

export function studioArtifacts(session) {
  const seen = new Map();
  const turns = session?.turns || [];
  for (let i = turns.length - 1; i >= 0; i -= 1) {
    const turn = turns[i];
    if (turn.file && ["pptx", "docx"].includes(turn.file.extension) && !seen.has(turn.file.id)) {
      seen.set(turn.file.id, {
        id: turn.file.id,
        kind: turn.file.extension,
        name: turn.file.name,
        file: turn.file
      });
    }
    if (Array.isArray(turn.images)) {
      for (let j = turn.images.length - 1; j >= 0; j -= 1) {
        const image = turn.images[j];
        const file = image?.file;
        if (file?.id && !seen.has(file.id)) {
          const title = (image.prompt || "生成图片").trim().slice(0, 26);
          seen.set(file.id, {
            id: file.id,
            kind: "image",
            name: title || "生成图片",
            file,
            image
          });
        }
      }
    }
  }
  return [...seen.values()];
}

export function selectionChip(selection) {
  return selection ? `<span class="attached-file">第 ${selection.slide} 页${selection.shape_id ? " · 选中文本框" : " · 整页文字"}
    <button type="button" data-act="slides-clear" aria-label="清除选区">×</button></span>` : "";
}

function artifactOptionLabel(item, list) {
  const sameKind = list.filter((entry) => entry.kind === item.kind);
  const idx = sameKind.indexOf(item);
  const prefix = item.kind === "pptx" ? "PPT · " : item.kind === "docx" ? "Word · " : "图片 · ";
  const suffix = sameKind.length > 1 ? (idx === 0 ? "（最新）" : `（旧版 ${idx}）`) : "";
  return `${prefix}${item.name}${suffix}`;
}

export function deckPanel(workspace) {
  const artifacts = studioArtifacts(workspace.session);
  const draft = readDraft(workspace.draftKey) || {};
  const activeItem = artifacts.find((item) => item.id === draft.deck?.fileId) || artifacts[0] || null;
  const hasItems = artifacts.length > 0;
  const dlHref = activeItem ? `/api/v1/files/${encodeURIComponent(activeItem.id)}/download` : "#";

  return `<section class="deck-studio" data-deck-preview aria-label="AI 实时创作画布">
    <div class="deck-toolbar">
      <div class="deck-file-wrap">
        <select class="deck-select" data-deck-file aria-label="切换画布成果" ${hasItems ? "" : "disabled"}>
          ${hasItems
            ? artifacts.map((item) => `<option value="${e(item.id)}" ${item.id === activeItem.id ? "selected" : ""}>${e(artifactOptionLabel(item, artifacts))}</option>`).join("")
            : '<option value="">等待 AI 生成成果…</option>'}
        </select>
      </div>
      <span class="live-pill" data-live-pill><i class="live-dot"></i><span data-live-pill-text>实时同步</span></span>
      <a class="button sm" href="${dlHref}" data-deck-download download ${hasItems ? "" : "hidden"}>${icon("download")}下载</a>
    </div>

    <div class="deck-stage" data-deck-stage>
      <div class="live-updating-banner" data-live-banner hidden>
        <span class="running-pulse" aria-hidden="true"></span>
        <span data-live-banner-text>AI 正在实时更新画布…</span>
      </div>

      <div class="deck-building" data-stage-building ${hasItems ? "hidden" : ""}>
        <div class="building-canvas-card" data-building-card>
          <div class="building-slide-skeleton" aria-hidden="true">
            <span class="skel-bar skel-tag"></span>
            <span class="skel-bar skel-title"></span>
            <span class="skel-bar skel-line"></span>
            <span class="skel-bar skel-line short"></span>
            <div class="skel-cols">
              <span class="skel-block"></span>
              <span class="skel-block"></span>
            </div>
          </div>
          <div class="building-status">
            <strong data-building-title>实时创作画布已就绪</strong>
            <p class="hint" data-building-note>在左侧交代制作 PPT、起草 Word 或生成图片，这里会同步展示过程与成稿。</p>
            <ol class="building-steps" data-building-steps hidden></ol>
          </div>
        </div>
      </div>

      <div class="deck-canvas-frame" data-stage-pptx hidden>
        <canvas data-deck-canvas role="img" aria-label="幻灯片预览图" hidden></canvas>
        <p class="deck-loading hint" data-deck-loading role="status">正在加载幻灯片…</p>
      </div>

      <div class="doc-stage-frame" data-stage-docx hidden>
        <article class="doc-paper">
          <header class="doc-paper-head">
            <strong data-doc-title></strong>
            <button type="button" class="text-button" data-act="edit-file" data-doc-edit-btn>改字</button>
          </header>
          <pre class="doc-paper-body" data-doc-body></pre>
        </article>
      </div>

      <div class="image-stage-frame" data-stage-image hidden>
        <button type="button" class="stage-image-btn" data-act="image-lightbox" data-stage-img-btn aria-label="放大查看图片">
          <img data-stage-img alt="">
          <span class="turn-image-zoom">${icon("maximize")}</span>
        </button>
        <p class="hint stage-image-meta" data-stage-img-meta></p>
      </div>
    </div>

    <div class="deck-pptx-footer" data-pptx-controls hidden>
      <div class="deck-strip" data-deck-strip aria-label="幻灯片页码快切"></div>
      <div class="deck-pager">
        <button type="button" class="icon-button bordered" data-deck-prev aria-label="上一页" title="上一页" disabled>${icon("left")}</button>
        <select class="deck-page-select" data-deck-page aria-label="切换幻灯片页码" disabled></select>
        <span class="deck-counter mono" data-deck-counter>— / —</span>
        <button type="button" class="icon-button bordered" data-deck-next aria-label="下一页" title="下一页" disabled>${icon("right")}</button>
      </div>
      <div class="deck-dock">
        <div class="deck-dock-row">
          <select class="deck-shape-select" data-deck-shape aria-label="选择修改范围" disabled>
            <option value="">整页的可编辑文字</option>
          </select>
          ${button("引用选区", "slides-target", { cls: "primary sm", disabled: true })}
        </div>
        <p class="hint deck-note" data-deck-note role="status">点击画布中任意文本框，自动选中并在左侧直接对话修改。</p>
      </div>
    </div>
  </section>`;
}

function textLines(ctx, text, maxWidth) {
  return text.replace(/\r/g, "").split(/\n|\v/).flatMap((line) => {
    const lines = [];
    let row = "";
    for (const char of line) {
      if (row && ctx.measureText(row + char).width > maxWidth) {
        lines.push(row);
        row = char;
      } else row += char;
    }
    return [...lines, row];
  });
}

function hexToRgba(hex, alpha = 1) {
  const clean = String(hex || "#000000").replace(/^#/, "");
  if (clean.length !== 6) return hex || "#000000";
  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${Math.max(0, Math.min(1, alpha))})`;
}

function traceShapePath(ctx, shape, left, top, width, height) {
  ctx.beginPath();
  if (shape.shape === "ellipse") {
    ctx.ellipse(left + width / 2, top + height / 2, Math.max(0.5, width / 2), Math.max(0.5, height / 2), 0, 0, Math.PI * 2);
  } else if (shape.shape === "round_rect") {
    const ratio = typeof shape.radius === "number" ? shape.radius : 0.08;
    const r = Math.min(28, Math.min(width, height) * Math.max(0.02, Math.min(0.5, ratio)));
    if (typeof ctx.roundRect === "function") {
      ctx.roundRect(left, top, width, height, r);
    } else {
      ctx.rect(left, top, width, height);
    }
  } else {
    ctx.rect(left, top, width, height);
  }
}

async function art(canvas, scene, selected) {
  const scale = 1.6;
  canvas.width = Math.round(scene.width * scale);
  canvas.height = Math.round(scene.height * scale);
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  const images = await Promise.all(scene.objects.map(async (shape) => {
    if (!shape.image) return null;
    const image = new Image();
    image.src = shape.image;
    try { await image.decode(); return image; } catch { return null; }
  }));
  ctx.fillStyle = scene.background;
  ctx.fillRect(0, 0, scene.width, scene.height);
  for (const [index, shape] of scene.objects.entries()) {
    const { left, top, width, height } = shape;
    if (shape.fill || shape.gradient || shape.border) {
      ctx.save();
      if (typeof shape.opacity === "number" && shape.opacity < 1) ctx.globalAlpha = Math.max(0, Math.min(1, shape.opacity));
      if (shape.shadow) {
        ctx.shadowColor = "rgba(15, 23, 42, 0.11)";
        ctx.shadowBlur = 14;
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 4;
      }
      traceShapePath(ctx, shape, left, top, width, height);
      if (shape.gradient?.stops?.length >= 2) {
        const rad = ((Number(shape.gradient.angle) || 0) * Math.PI) / 180;
        const cx = left + width / 2;
        const cy = top + height / 2;
        const dx = Math.cos(rad) * (width / 2);
        const dy = Math.sin(rad) * (height / 2);
        const grad = ctx.createLinearGradient(cx - dx, cy - dy, cx + dx, cy + dy);
        for (const stop of shape.gradient.stops) {
          grad.addColorStop(Math.max(0, Math.min(1, stop.pos)), hexToRgba(stop.color, stop.alpha));
        }
        ctx.fillStyle = grad;
        ctx.fill();
      } else if (shape.fill) {
        ctx.fillStyle = shape.fill;
        ctx.fill();
      }
      if (shape.border) {
        ctx.shadowColor = "transparent";
        ctx.strokeStyle = shape.border;
        ctx.lineWidth = 1;
        ctx.stroke();
      }
      ctx.restore();
    }
    if (images[index]) {
      const img = images[index];
      // 按 cover 等比绘制预览，避免拉伸变形
      const ir = img.naturalWidth / Math.max(1, img.naturalHeight);
      const br = width / Math.max(1, height);
      let sx = 0, sy = 0, sw = img.naturalWidth, sh = img.naturalHeight;
      if (ir > br) {
        sw = img.naturalHeight * br;
        sx = (img.naturalWidth - sw) / 2;
      } else if (ir < br) {
        sh = img.naturalWidth / br;
        sy = (img.naturalHeight - sh) / 2;
      }
      ctx.drawImage(img, sx, sy, sw, sh, left, top, width, height);
    }
    if (shape.text) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(left, top, width, height);
      ctx.clip();
      ctx.font = `${shape.bold ? "700 " : ""}${Math.max(7, shape.size)}px "Microsoft YaHei","Noto Sans CJK SC",sans-serif`;
      ctx.fillStyle = shape.color;
      ctx.textBaseline = "top";
      const align = shape.align === "center" ? "center" : shape.align === "right" ? "right" : "left";
      ctx.textAlign = align;
      const lines = textLines(ctx, shape.text, Math.max(10, width - 4)).slice(0, 100);
      const lineH = shape.size * 1.24;
      const totalH = lines.length * lineH;
      const startY = shape.valign === "middle"
        ? top + Math.max(1, (height - totalH) / 2)
        : shape.valign === "bottom"
          ? top + Math.max(1, height - totalH - 2)
          : top + 2;
      const drawX = align === "center" ? left + width / 2 : align === "right" ? left + width - 2 : left + 2;
      lines.forEach((line, lineIndex) => ctx.fillText(line, drawX, startY + lineIndex * lineH));
      ctx.restore();
    }
    if (shape.id === selected) {
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 4;
      ctx.strokeRect(left, top, width, height);
      ctx.strokeStyle = "#18181b";
      ctx.lineWidth = 2;
      ctx.setLineDash([7, 4]);
      ctx.strokeRect(left, top, width, height);
      ctx.setLineDash([]);
    }
  }
}

function remember(workspace, fields) {
  saveDraft(workspace.draftKey, { ...(readDraft(workspace.draftKey) || {}), ...fields });
}

export function clearSelection() {
  const workspace = state.chatWorkspace;
  if (!workspace) return;
  remember(workspace, { selection: null });
  const chip = document.querySelector("[data-chat-selection]");
  if (chip) chip.innerHTML = "";
}

function rebuildChoices(select, options) {
  select.replaceChildren(...options.map(({ text, value }) => new Option(text, value)));
}

let activeController = null;
let liveUpdatedTimer = null;

export function selectDeckFile(fileId, preferredPage = 1) {
  if (activeController && fileId) {
    void activeController.selectById(fileId, preferredPage);
    return;
  }
  const panel = document.querySelector("[data-deck-preview]");
  const file = panel?.querySelector("[data-deck-file]");
  if (!file || !fileId) return;
  if ([...file.options].some((opt) => opt.value === fileId) && file.value !== fileId) {
    file.value = fileId;
    file.dispatchEvent(new Event("change"));
  }
}

export function setCanvasLiveState({ active = false, kind = "pptx", label = "", prompt = "", steps = [], targetPage = null } = {}) {
  const panel = document.querySelector("[data-deck-preview]");
  if (!panel) return;
  const pill = panel.querySelector("[data-live-pill]");
  const pillText = panel.querySelector("[data-live-pill-text]");
  const banner = panel.querySelector("[data-live-banner]");
  const bannerText = panel.querySelector("[data-live-banner-text]");
  const stage = panel.querySelector("[data-deck-stage]");
  const building = panel.querySelector("[data-stage-building]");
  const buildCard = panel.querySelector("[data-building-card]");
  const buildTitle = panel.querySelector("[data-building-title]");
  const buildNote = panel.querySelector("[data-building-note]");
  const buildSteps = panel.querySelector("[data-building-steps]");
  const hasExisting = Boolean(panel.querySelector("[data-deck-file]")?.value);

  clearTimeout(liveUpdatedTimer);
  stage?.classList.toggle("is-live-updating", Boolean(active));

  if (active) {
    const kindName = kind === "image" ? "图片" : kind === "docx" ? "Word 文稿" : "PPT 演示文稿";
    const statusText = label || (targetPage ? `正在实时修改第 ${targetPage} 页…` : `正在实时制作${kindName}…`);
    if (pill) {
      pill.classList.remove("updated");
      pill.classList.add("working");
    }
    if (pillText) pillText.textContent = "实时生成中";

    if (hasExisting) {
      if (banner) banner.hidden = false;
      if (bannerText) bannerText.textContent = statusText;
    } else {
      if (banner) banner.hidden = true;
      if (building) building.hidden = false;
      if (buildCard) buildCard.classList.add("is-building");
      if (buildTitle) buildTitle.textContent = targetPage ? `正在更新第 ${targetPage} 页` : `AI 正在制作${kindName}`;
      if (buildNote) buildNote.textContent = label || (prompt ? `目标：${prompt.slice(0, 60)}` : "正在实时规划内容与排版…");
      if (buildSteps) {
        if (steps.length) {
          buildSteps.hidden = false;
          buildSteps.innerHTML = steps.map((text, idx) =>
            `<li class="${idx === steps.length - 1 ? "active" : "done"}"><span class="step-dot"></span><span>${e(text)}</span></li>`
          ).join("");
        } else {
          buildSteps.hidden = true;
        }
      }
    }
  } else {
    if (banner) banner.hidden = true;
    if (buildCard) buildCard.classList.remove("is-building");
    if (pill) pill.classList.remove("working");
    if (pillText) pillText.textContent = "实时同步";
  }
}

function flashUpdated(panel) {
  const pill = panel.querySelector("[data-live-pill]");
  const pillText = panel.querySelector("[data-live-pill-text]");
  const stage = panel.querySelector("[data-deck-stage]");
  clearTimeout(liveUpdatedTimer);
  if (pill) {
    pill.classList.remove("working");
    pill.classList.add("updated");
  }
  if (pillText) pillText.textContent = "已实时更新";
  if (stage) {
    stage.classList.remove("just-updated");
    void stage.offsetWidth;
    stage.classList.add("just-updated");
  }
  liveUpdatedTimer = setTimeout(() => {
    if (pill?.isConnected) pill.classList.remove("updated");
    if (pillText?.isConnected) pillText.textContent = "实时同步";
    if (stage?.isConnected) stage.classList.remove("just-updated");
  }, 2600);
}

export async function refreshCanvasLive(workspace, { selectId = "", preferredPage = 1, flash = true } = {}) {
  const panel = document.querySelector("[data-deck-preview]");
  if (!panel || !workspace) return;
  if (!activeController || !activeController.isMountedOn(panel)) {
    mountDeck();
  }
  await activeController?.syncArtifacts(workspace, { selectId, preferredPage, flash });
}

export function mountDeck() {
  const panel = document.querySelector("[data-deck-preview]");
  const workspace = state.chatWorkspace;
  if (!panel || !workspace) {
    activeController = null;
    return;
  }
  const file = panel.querySelector("[data-deck-file]");
  const dlLink = panel.querySelector("[data-deck-download]");
  const stageBuilding = panel.querySelector("[data-stage-building]");
  const stagePptx = panel.querySelector("[data-stage-pptx]");
  const stageDocx = panel.querySelector("[data-stage-docx]");
  const stageImage = panel.querySelector("[data-stage-image]");
  const pptxControls = panel.querySelector("[data-pptx-controls]");
  const strip = panel.querySelector("[data-deck-strip]");
  const page = panel.querySelector("[data-deck-page]");
  const prevBtn = panel.querySelector("[data-deck-prev]");
  const nextBtn = panel.querySelector("[data-deck-next]");
  const counter = panel.querySelector("[data-deck-counter]");
  const shape = panel.querySelector("[data-deck-shape]");
  const canvas = panel.querySelector("canvas");
  const loading = panel.querySelector("[data-deck-loading]");
  const note = panel.querySelector("[data-deck-note]");
  const action = panel.querySelector("[data-act=slides-target]");
  const docTitle = panel.querySelector("[data-doc-title]");
  const docBody = panel.querySelector("[data-doc-body]");
  const docEditBtn = panel.querySelector("[data-doc-edit-btn]");
  const imgBtn = panel.querySelector("[data-stage-img-btn]");
  const imgEl = panel.querySelector("[data-stage-img]");
  const imgMeta = panel.querySelector("[data-stage-img-meta]");

  let currentWorkspace = workspace;
  let artifacts = studioArtifacts(currentWorkspace.session);
  let scene = null;
  let epoch = 0;
  let totalPages = 0;

  function setStageMode(mode) {
    if (stageBuilding) stageBuilding.hidden = mode !== "building";
    if (stagePptx) stagePptx.hidden = mode !== "pptx";
    if (stageDocx) stageDocx.hidden = mode !== "docx";
    if (stageImage) stageImage.hidden = mode !== "image";
    if (pptxControls) pptxControls.hidden = mode !== "pptx";
  }

  function currentArtifact() {
    return artifacts.find((item) => item.id === file.value) || null;
  }

  function syncPager(current) {
    if (counter) {
      counter.textContent = totalPages ? `${String(current).padStart(2, "0")} / ${String(totalPages).padStart(2, "0")}` : "— / —";
    }
    if (prevBtn) prevBtn.disabled = !totalPages || current <= 1;
    if (nextBtn) nextBtn.disabled = !totalPages || current >= totalPages;
    if (strip) {
      for (const pill of strip.querySelectorAll("[data-strip-page]")) {
        const isCurrent = Number(pill.dataset.stripPage) === current;
        pill.classList.toggle("active", isCurrent);
        pill.setAttribute("aria-current", isCurrent ? "true" : "false");
      }
    }
  }

  async function loadPptxPage(fileId, number) {
    const ticket = ++epoch;
    scene = null;
    action.disabled = true;
    syncPager(number);
    if (loading) { loading.textContent = "正在渲染幻灯片…"; loading.hidden = false; }
    const next = await api("/files/" + encodeURIComponent(fileId) + "/slides/" + number);
    if (!panel.isConnected || ticket !== epoch) return;
    scene = next;
    const objects = scene.objects.filter((item) => item.editable);
    rebuildChoices(shape, [{ text: "整页的可编辑文字", value: "" }, ...objects.map((item) => ({
      text: item.text.slice(0, 31), value: String(item.id)
    }))]);
    const selected = (readDraft(currentWorkspace.draftKey) || {}).selection;
    if (selected?.file_id === fileId && selected.slide === number && selected.shape_id) {
      shape.value = String(selected.shape_id);
    }
    await art(canvas, next, Number(shape.value));
    if (!panel.isConnected || ticket !== epoch) return;
    canvas.hidden = false;
    canvas.classList.remove("fading");
    void canvas.offsetWidth;
    canvas.classList.add("fading");
    if (loading) loading.hidden = true;
    canvas.setAttribute("aria-label", `第 ${number} 页预览`);
    shape.disabled = !objects.length;
    action.disabled = !(objects.length && document.querySelector("[data-form=conversation]"));
    note.textContent = objects.length ? "点击画布任意文本框，直接选中并在左侧对话修改。" : "当前页没有可编辑的单段文本框。";
    remember(currentWorkspace, { deck: { fileId, page: number } });
  }

  async function loadPptxFile(fileId, preferred = 1) {
    const ticket = ++epoch;
    setStageMode("pptx");
    scene = null;
    totalPages = 0;
    page.disabled = true;
    shape.disabled = true;
    action.disabled = true;
    canvas.hidden = true;
    if (strip) strip.replaceChildren();
    syncPager(0);
    if (loading) { loading.textContent = "正在加载演示文稿…"; loading.hidden = false; }
    const result = await api("/files/" + encodeURIComponent(fileId) + "/slides");
    if (!panel.isConnected || ticket !== epoch) return;
    totalPages = result.pages.length;
    rebuildChoices(page, result.pages.map((slide) => ({ text: `${slide.page}. ${slide.title}`, value: String(slide.page) })));
    if (strip) {
      strip.innerHTML = result.pages.map((slide) =>
        `<button type="button" class="deck-strip-btn" data-strip-page="${slide.page}" title="${e(`${slide.page}. ${slide.title}`)}">${String(slide.page).padStart(2, "0")}</button>`
      ).join("");
    }
    const targetPage = Math.max(1, Math.min(preferred, totalPages || 1));
    page.value = String(targetPage);
    page.disabled = totalPages <= 1;
    await loadPptxPage(fileId, targetPage);
  }

  async function loadDocxFile(item) {
    const ticket = ++epoch;
    setStageMode("docx");
    if (docTitle) docTitle.textContent = item.name;
    if (docEditBtn) docEditBtn.dataset.id = item.id;
    if (docBody) docBody.textContent = "正在加载文稿内容…";
    const preview = await api("/files/" + encodeURIComponent(item.id) + "/preview");
    if (!panel.isConnected || ticket !== epoch) return;
    if (docTitle) docTitle.textContent = preview.name || item.name;
    if (docBody) docBody.textContent = preview.text || "此文稿已生成，可点击上方下载查看完整格式。";
    remember(currentWorkspace, { deck: { fileId: item.id, page: 1 } });
  }

  function loadImageItem(item) {
    ++epoch;
    setStageMode("image");
    const image = item.image || {};
    const quality = image.quality === "low" ? "草稿" : image.quality === "high" ? "精细" : "标准";
    const meta = image.width && image.height ? `${image.width}×${image.height} · ${quality}质量 · ${image.model || ""}` : item.name;
    const alt = image.prompt || item.name || "生成的图片";
    if (imgEl) {
      imgEl.src = `/api/v1/files/${encodeURIComponent(item.id)}/image`;
      imgEl.alt = alt;
    }
    if (imgBtn) {
      imgBtn.dataset.id = item.id;
      imgBtn.dataset.alt = alt;
      imgBtn.dataset.meta = meta;
    }
    if (imgMeta) imgMeta.textContent = meta;
    remember(currentWorkspace, { deck: { fileId: item.id, page: 1 } });
  }

  async function loadSelectedArtifact(preferredPage = 1) {
    const item = currentArtifact();
    if (!item) {
      setStageMode("building");
      if (dlLink) dlLink.hidden = true;
      return;
    }
    if (dlLink) {
      dlLink.href = "/api/v1/files/" + encodeURIComponent(item.id) + "/download";
      dlLink.hidden = false;
    }
    if (item.kind === "pptx") await loadPptxFile(item.id, preferredPage);
    else if (item.kind === "docx") await loadDocxFile(item);
    else if (item.kind === "image") loadImageItem(item);
  }

  function report(error) {
    if (panel.isConnected) {
      if (loading) loading.hidden = true;
      if (note) note.textContent = error.message;
    }
  }

  function applySelection(focusInput = true) {
    const form = document.querySelector("[data-form=conversation]");
    if (!form || form.dataset.busy || !scene) return;
    const selection = { file_id: file.value, slide: Number(page.value), shape_id: shape.value ? Number(shape.value) : null };
    remember(currentWorkspace, { selection });
    form.querySelector("[data-chat-selection]").innerHTML = selectionChip(selection);
    form.classList.remove("pulsing");
    void form.offsetWidth;
    form.classList.add("pulsing");
    if (focusInput) form.elements.message.focus();
  }

  file.addEventListener("change", () => { clearSelection(); void loadSelectedArtifact(1).catch(report); });
  page.addEventListener("change", () => {
    const item = currentArtifact();
    if (item?.kind === "pptx") { clearSelection(); void loadPptxPage(item.id, Number(page.value)).catch(report); }
  });
  strip?.addEventListener("click", (event) => {
    const btn = event.target.closest("[data-strip-page]");
    const item = currentArtifact();
    if (!btn || item?.kind !== "pptx") return;
    const num = Number(btn.dataset.stripPage);
    if (num && num !== Number(page.value)) {
      page.value = String(num);
      clearSelection();
      void loadPptxPage(item.id, num).catch(report);
    }
  });
  prevBtn?.addEventListener("click", () => {
    const item = currentArtifact();
    const cur = Number(page.value);
    if (item?.kind === "pptx" && cur > 1) {
      page.value = String(cur - 1);
      clearSelection();
      void loadPptxPage(item.id, cur - 1).catch(report);
    }
  });
  nextBtn?.addEventListener("click", () => {
    const item = currentArtifact();
    const cur = Number(page.value);
    if (item?.kind === "pptx" && cur < totalPages) {
      page.value = String(cur + 1);
      clearSelection();
      void loadPptxPage(item.id, cur + 1).catch(report);
    }
  });
  shape.addEventListener("change", () => {
    clearSelection();
    if (scene) void art(canvas, scene, Number(shape.value));
  });
  canvas.addEventListener("click", (event) => {
    if (!scene || shape.disabled) return;
    const bounds = canvas.getBoundingClientRect();
    const x = (event.clientX - bounds.left) * scene.width / bounds.width;
    const y = (event.clientY - bounds.top) * scene.height / bounds.height;
    const hit = [...scene.objects].reverse().find((item) => item.editable && x >= item.left && x <= item.left + item.width && y >= item.top && y <= item.top + item.height);
    if (hit) {
      shape.value = String(hit.id);
      void art(canvas, scene, hit.id);
      applySelection(true);
      toastMessage(`已选中第 ${page.value} 页文本框，可在左侧直接输入修改要求`, "info");
    }
  });
  action.addEventListener("click", () => {
    applySelection(true);
    toastMessage("已选中要修改的幻灯片范围", "success");
  });

  activeController = {
    isMountedOn: (el) => el === panel && panel.isConnected,
    async selectById(fileId, preferredPage = 1) {
      if (!artifacts.some((item) => item.id === fileId)) return;
      file.value = fileId;
      await loadSelectedArtifact(preferredPage).catch(report);
    },
    async syncArtifacts(nextWorkspace, { selectId = "", preferredPage = 1, flash = true } = {}) {
      currentWorkspace = nextWorkspace;
      artifacts = studioArtifacts(currentWorkspace.session);
      const workspaceEl = document.querySelector(".agent-workspace");
      if (workspaceEl) workspaceEl.dataset.hasDeck = artifacts.length ? "1" : "0";
      if (!artifacts.length) {
        rebuildChoices(file, [{ text: "等待 AI 生成成果…", value: "" }]);
        file.disabled = true;
        setStageMode("building");
        if (dlLink) dlLink.hidden = true;
        return;
      }
      const targetId = selectId && artifacts.some((item) => item.id === selectId)
        ? selectId
        : artifacts[0].id;
      rebuildChoices(file, artifacts.map((item) => ({ text: artifactOptionLabel(item, artifacts), value: item.id })));
      file.disabled = false;
      file.value = targetId;
      setCanvasLiveState({ active: false });
      await loadSelectedArtifact(preferredPage).catch(report);
      if (flash && panel.isConnected) flashUpdated(panel);
    }
  };

  const saved = readDraft(currentWorkspace.draftKey) || {};
  void loadSelectedArtifact(saved.deck?.page || 1).catch(report);
}
