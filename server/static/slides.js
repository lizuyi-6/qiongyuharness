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

export function selectionSummaryText(selection) {
  if (!selection) return "";
  const pageText = `第 ${selection.slide} 页`;
  if (selection.label) return `${pageText} · ${selection.label}`;
  if (selection.region) return `${pageText} · 圈选区域`;
  if (Array.isArray(selection.shape_ids) && selection.shape_ids.length > 1) {
    return `${pageText} · 选中 ${selection.shape_ids.length} 个元素`;
  }
  if (selection.shape_id) return `${pageText} · 选中元素`;
  return `${pageText} · 整页自由修改`;
}

export function selectionChip(selection) {
  return selection ? `<span class="attached-file">${e(selectionSummaryText(selection))}
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
            <option value="">整页自由修改（含图片、排版与文字）</option>
          </select>
          <button type="button" class="button sm ghost" data-deck-clear-box hidden>清除框选</button>
          ${button("引用选区", "slides-target", { cls: "primary sm", disabled: true })}
        </div>
        <p class="hint deck-note" data-deck-note role="status">在画布上拖拽鼠标可自由圈选任意区域（含图片、卡片、文字），或单击任意元素直接让 AI 修改。</p>
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

const SLIDE_W_IN = 13.333;
const SLIDE_H_IN = 7.5;

function boxesIntersect(a, b) {
  return a.left < b.left + b.width && a.left + a.width > b.left && a.top < b.top + b.height && a.top + a.height > b.top;
}

function normalizeBox(p1, p2, maxW, maxH) {
  const left = Math.max(0, Math.min(maxW, Math.min(p1.x, p2.x)));
  const top = Math.max(0, Math.min(maxH, Math.min(p1.y, p2.y)));
  const right = Math.max(0, Math.min(maxW, Math.max(p1.x, p2.x)));
  const bottom = Math.max(0, Math.min(maxH, Math.max(p1.y, p2.y)));
  return { left, top, width: Math.max(1, right - left), height: Math.max(1, bottom - top) };
}

function objectKind(item) {
  if (item?.kind) return item.kind;
  if (item?.image) return "image";
  if (item?.text) return "text";
  return "shape";
}

function objectOptionLabel(item, counters) {
  const kind = objectKind(item);
  if (kind === "image") {
    counters.image += 1;
    return item.is_background ? `[图片] 背景主图` : `[图片] 配图 #${counters.image}`;
  }
  if (kind === "text") {
    const snippet = String(item.text || "").replace(/\s+/g, " ").trim().slice(0, 28);
    return `[文字] ${snippet || "文本框"}`;
  }
  counters.shape += 1;
  return `[图形] 卡片/色块 #${counters.shape}`;
}

async function renderSceneBase(scene) {
  const scale = 1.6;
  const offscreen = document.createElement("canvas");
  offscreen.width = Math.round(scene.width * scale);
  offscreen.height = Math.round(scene.height * scale);
  const ctx = offscreen.getContext("2d");
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
  }
  return { offscreen, scale };
}

function paintOverlay(canvas, base, scene, { selectedIds = [], regionBox = null, isDragging = false } = {}) {
  if (!base || !scene) return;
  const { offscreen, scale } = base;
  if (canvas.width !== offscreen.width || canvas.height !== offscreen.height) {
    canvas.width = offscreen.width;
    canvas.height = offscreen.height;
  }
  const ctx = canvas.getContext("2d");
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(offscreen, 0, 0);
  ctx.scale(scale, scale);

  const selectedSet = new Set(selectedIds.filter(Boolean));
  for (const shape of scene.objects) {
    if (!selectedSet.has(shape.id)) continue;
    const { left, top, width, height } = shape;
    ctx.save();
    ctx.strokeStyle = "rgba(255,255,255,0.9)";
    ctx.lineWidth = regionBox ? 2.5 : 4;
    ctx.strokeRect(left, top, width, height);
    ctx.strokeStyle = regionBox ? "rgba(37, 99, 235, 0.72)" : "#2563eb";
    ctx.lineWidth = regionBox ? 1.5 : 2.2;
    ctx.setLineDash([6, 4]);
    ctx.strokeRect(left, top, width, height);
    ctx.setLineDash([]);
    ctx.restore();
  }

  if (regionBox) {
    const { left, top, width, height } = regionBox;
    ctx.save();
    ctx.fillStyle = isDragging ? "rgba(37, 99, 235, 0.14)" : "rgba(37, 99, 235, 0.10)";
    ctx.fillRect(left, top, width, height);
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 3.5;
    ctx.strokeRect(left, top, width, height);
    ctx.strokeStyle = "#2563eb";
    ctx.lineWidth = 2;
    ctx.setLineDash(isDragging ? [6, 4] : []);
    ctx.strokeRect(left, top, width, height);
    ctx.setLineDash([]);
    // 四角控制锚点
    const r = 4;
    for (const [hx, hy] of [[left, top], [left + width, top], [left, top + height], [left + width, top + height]]) {
      ctx.fillStyle = "#ffffff";
      ctx.strokeStyle = "#2563eb";
      ctx.lineWidth = 1.6;
      ctx.fillRect(hx - r, hy - r, r * 2, r * 2);
      ctx.strokeRect(hx - r, hy - r, r * 2, r * 2);
    }
    const wIn = (width / scene.width * SLIDE_W_IN).toFixed(1);
    const hIn = (height / scene.height * SLIDE_H_IN).toFixed(1);
    const tag = isDragging ? `圈选区域 ${wIn}" × ${hIn}"` : `已圈选 · ${wIn}" × ${hIn}"`;
    ctx.font = '600 11px "Microsoft YaHei", sans-serif';
    const tw = ctx.measureText(tag).width + 12;
    const th = 20;
    const tx = Math.max(4, Math.min(scene.width - tw - 4, left));
    const ty = top >= th + 6 ? top - th - 4 : top + 6;
    ctx.fillStyle = "#2563eb";
    ctx.fillRect(tx, ty, tw, th);
    ctx.fillStyle = "#ffffff";
    ctx.textBaseline = "middle";
    ctx.textAlign = "left";
    ctx.fillText(tag, tx + 6, ty + th / 2);
    ctx.restore();
  }
  ctx.restore();
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
  activeController?.clearBox?.();
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
  const clearBoxBtn = panel.querySelector("[data-deck-clear-box]");
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
  let baseLayer = null;
  let epoch = 0;
  let totalPages = 0;
  let activeRegion = null; // { left, top, width, height } in scene coordinates
  let activeShapeIds = [];
  let activeLabel = "";
  let dragStart = null;
  let dragging = false;

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

  function redrawCanvas(isDragging = false) {
    if (!baseLayer || !scene) return;
    const selectedIds = activeRegion
      ? activeShapeIds
      : (shape.value ? [Number(shape.value)] : []);
    paintOverlay(canvas, baseLayer, scene, { selectedIds, regionBox: activeRegion, isDragging });
    if (clearBoxBtn) clearBoxBtn.hidden = !activeRegion && !shape.value;
  }

  function describeRegionHits(box) {
    if (!scene || !box) return { ids: [], label: "圈选区域" };
    const fgHits = [];
    let bgHit = null;
    for (const item of scene.objects) {
      if (!boxesIntersect(item, box)) continue;
      if (item.is_background) {
        bgHit = item;
        continue;
      }
      fgHits.push(item);
    }
    const ids = fgHits.map((item) => item.id);
    let images = fgHits.filter((item) => objectKind(item) === "image").length;
    const texts = fgHits.filter((item) => objectKind(item) === "text").length;
    const shapes = fgHits.filter((item) => objectKind(item) === "shape").length;
    if (bgHit && (!fgHits.length || (box.width * box.height) / (scene.width * scene.height) >= 0.22)) {
      ids.push(bgHit.id);
      images += 1;
    }
    const parts = [];
    if (images) parts.push(`${images} 张图片`);
    if (texts) parts.push(`${texts} 处文字`);
    if (shapes) parts.push(`${shapes} 个图形`);
    const label = parts.length ? `圈选区域（含 ${parts.join(" · ")}）` : "圈选空白区域（可新增图文）";
    return { ids, label };
  }

  function describeSingleObject(item) {
    if (!item) return "整页自由修改";
    const kind = objectKind(item);
    if (kind === "image") return item.is_background ? "背景主图" : "选中图片";
    if (kind === "text") {
      const snippet = String(item.text || "").replace(/\s+/g, " ").trim().slice(0, 14);
      return snippet ? `文字「${snippet}」` : "选中文本框";
    }
    return "选中卡片/图形";
  }

  function scenePointFromEvent(event) {
    const bounds = canvas.getBoundingClientRect();
    const x = Math.max(0, Math.min(scene.width, (event.clientX - bounds.left) * scene.width / Math.max(1, bounds.width)));
    const y = Math.max(0, Math.min(scene.height, (event.clientY - bounds.top) * scene.height / Math.max(1, bounds.height)));
    return { x, y };
  }

  function pickObjectAtPoint(x, y) {
    if (!scene) return null;
    const hits = scene.objects.filter((item) =>
      x >= item.left && x <= item.left + item.width && y >= item.top && y <= item.top + item.height
    );
    if (!hits.length) return null;
    const foreground = hits.filter((item) => !item.is_background);
    if (foreground.length) {
      // 优先命中上层文字/图片，面积更小者优先，方便在卡片内部直接点中图片或文本
      const score = (item) => {
        const kind = objectKind(item);
        const priority = kind === "text" ? 0 : kind === "image" ? 1 : 2;
        return priority * 1e9 + (item.width * item.height);
      };
      foreground.sort((a, b) => score(a) - score(b));
      return foreground[0];
    }
    return hits[hits.length - 1];
  }

  async function loadPptxPage(fileId, number) {
    const ticket = ++epoch;
    scene = null;
    baseLayer = null;
    activeRegion = null;
    activeShapeIds = [];
    activeLabel = "";
    action.disabled = true;
    syncPager(number);
    if (loading) { loading.textContent = "正在渲染幻灯片…"; loading.hidden = false; }
    const next = await api("/files/" + encodeURIComponent(fileId) + "/slides/" + number);
    if (!panel.isConnected || ticket !== epoch) return;
    scene = next;
    const counters = { image: 0, shape: 0 };
    const selectableObjects = scene.objects.filter((item) => item.selectable !== false);
    rebuildChoices(shape, [
      { text: "整页自由修改（含图片、排版与文字）", value: "" },
      ...selectableObjects.map((item) => ({
        text: objectOptionLabel(item, counters),
        value: String(item.id)
      }))
    ]);
    const selected = (readDraft(currentWorkspace.draftKey) || {}).selection;
    if (selected?.file_id === fileId && selected.slide === number) {
      if (selected.region && Number.isFinite(selected.region.w)) {
        activeRegion = {
          left: selected.region.x / SLIDE_W_IN * scene.width,
          top: selected.region.y / SLIDE_H_IN * scene.height,
          width: selected.region.w / SLIDE_W_IN * scene.width,
          height: selected.region.h / SLIDE_H_IN * scene.height
        };
        activeShapeIds = Array.isArray(selected.shape_ids) ? selected.shape_ids : [];
        activeLabel = selected.label || "圈选区域";
      } else if (selected.shape_id) {
        shape.value = String(selected.shape_id);
        activeShapeIds = [Number(selected.shape_id)];
        activeLabel = selected.label || "";
      }
    }
    baseLayer = await renderSceneBase(next);
    if (!panel.isConnected || ticket !== epoch) return;
    redrawCanvas(false);
    canvas.hidden = false;
    canvas.classList.remove("fading");
    void canvas.offsetWidth;
    canvas.classList.add("fading");
    if (loading) loading.hidden = true;
    canvas.setAttribute("aria-label", `第 ${number} 页预览（支持拖拽框选或点击任意元素）`);
    shape.disabled = false;
    action.disabled = !document.querySelector("[data-form=conversation]");
    note.textContent = activeRegion
      ? `已${activeLabel || "圈选区域"}，可在左侧直接输入修改要求（支持换图、改字、调排版）。`
      : "在画布上拖拽鼠标可自由圈选任意区域（含图片、卡片、文字），或单击任意元素直接让 AI 修改。";
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
    let selection;
    if (activeRegion) {
      const region = {
        x: Number((activeRegion.left / scene.width * SLIDE_W_IN).toFixed(2)),
        y: Number((activeRegion.top / scene.height * SLIDE_H_IN).toFixed(2)),
        w: Number(Math.max(0.1, activeRegion.width / scene.width * SLIDE_W_IN).toFixed(2)),
        h: Number(Math.max(0.1, activeRegion.height / scene.height * SLIDE_H_IN).toFixed(2))
      };
      selection = {
        file_id: file.value,
        slide: Number(page.value),
        shape_id: activeShapeIds.length === 1 ? activeShapeIds[0] : null,
        shape_ids: activeShapeIds.slice(0, 50),
        region,
        label: activeLabel || "圈选区域"
      };
    } else if (shape.value) {
      const sid = Number(shape.value);
      const obj = scene.objects.find((item) => item.id === sid);
      selection = {
        file_id: file.value,
        slide: Number(page.value),
        shape_id: sid,
        shape_ids: [sid],
        label: activeLabel || describeSingleObject(obj)
      };
    } else {
      selection = {
        file_id: file.value,
        slide: Number(page.value),
        shape_id: null,
        shape_ids: [],
        label: "整页自由修改"
      };
    }
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
    activeRegion = null;
    const sid = shape.value ? Number(shape.value) : null;
    activeShapeIds = sid ? [sid] : [];
    const obj = sid && scene ? scene.objects.find((item) => item.id === sid) : null;
    activeLabel = obj ? describeSingleObject(obj) : "";
    redrawCanvas(false);
    if (sid) applySelection(false);
    else clearSelection();
  });
  clearBoxBtn?.addEventListener("click", () => {
    activeRegion = null;
    activeShapeIds = [];
    activeLabel = "";
    shape.value = "";
    clearSelection();
    redrawCanvas(false);
    note.textContent = "在画布上拖拽鼠标可自由圈选任意区域（含图片、卡片、文字），或单击任意元素直接让 AI 修改。";
  });

  canvas.addEventListener("pointerdown", (event) => {
    if (!scene || !baseLayer || event.button !== 0) return;
    event.preventDefault();
    dragStart = scenePointFromEvent(event);
    dragging = false;
    try { canvas.setPointerCapture(event.pointerId); } catch {}
  });

  canvas.addEventListener("pointermove", (event) => {
    if (!dragStart || !scene || !baseLayer) return;
    const current = scenePointFromEvent(event);
    const dist = Math.hypot(current.x - dragStart.x, current.y - dragStart.y);
    if (!dragging && dist < 8) return;
    dragging = true;
    const box = normalizeBox(dragStart, current, scene.width, scene.height);
    activeRegion = box;
    const hitInfo = describeRegionHits(box);
    activeShapeIds = hitInfo.ids;
    activeLabel = hitInfo.label;
    shape.value = "";
    redrawCanvas(true);
  });

  function finishPointer(event, cancelled = false) {
    if (!dragStart || !scene || !baseLayer) return;
    const start = dragStart;
    const wasDragging = dragging;
    dragStart = null;
    dragging = false;
    try { canvas.releasePointerCapture(event.pointerId); } catch {}
    if (cancelled) {
      redrawCanvas(false);
      return;
    }
    const end = scenePointFromEvent(event);
    if (wasDragging && Math.abs(end.x - start.x) >= 10 && Math.abs(end.y - start.y) >= 10) {
      const box = normalizeBox(start, end, scene.width, scene.height);
      activeRegion = box;
      const hitInfo = describeRegionHits(box);
      activeShapeIds = hitInfo.ids;
      activeLabel = hitInfo.label;
      shape.value = "";
      redrawCanvas(false);
      applySelection(true);
      note.textContent = `已${activeLabel}，可在左侧直接让 AI 修改（支持换图、重绘、改排版或改文字）。`;
      toastMessage(`已圈选第 ${page.value} 页区域，可在左侧直接描述修改要求`, "info");
      return;
    }
    // 单击：命中任意元素（图片、文字、卡片/图形）
    activeRegion = null;
    const hit = pickObjectAtPoint(end.x, end.y);
    if (hit) {
      shape.value = String(hit.id);
      activeShapeIds = [hit.id];
      activeLabel = describeSingleObject(hit);
      redrawCanvas(false);
      applySelection(true);
      note.textContent = `已选中「${activeLabel}」，可在左侧直接输入修改要求。`;
      toastMessage(`已选中第 ${page.value} 页 ${activeLabel}`, "info");
    } else {
      shape.value = "";
      activeShapeIds = [];
      activeLabel = "";
      redrawCanvas(false);
    }
  }

  canvas.addEventListener("pointerup", (event) => finishPointer(event, false));
  canvas.addEventListener("pointercancel", (event) => finishPointer(event, true));

  action.addEventListener("click", () => {
    applySelection(true);
    toastMessage("已引用当前画布选区，可在左侧自由描述修改要求", "success");
  });

  activeController = {
    isMountedOn: (el) => el === panel && panel.isConnected,
    clearBox() {
      activeRegion = null;
      activeShapeIds = [];
      activeLabel = "";
      shape.value = "";
      redrawCanvas(false);
    },
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
