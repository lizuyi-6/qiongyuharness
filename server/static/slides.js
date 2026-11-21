import { api, button, e, icon, post, readDraft, saveDraft, state, toastMessage } from "./ui.js";

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
      <div class="deck-mode-bar" data-deck-mode-bar>
        <div class="deck-mode-switch" role="tablist" aria-label="画布交互模式">
          <button type="button" class="deck-mode-btn active" data-deck-mode="select" role="tab" aria-selected="true">${icon("crop")}框选给 AI</button>
          <button type="button" class="deck-mode-btn" data-deck-mode="move" role="tab" aria-selected="false">${icon("move")}拖动排版</button>
        </div>
        <div class="deck-layout-actions" data-deck-layout-actions hidden>
          <button type="button" class="button sm ghost" data-deck-reset-layout>还原</button>
          <button type="button" class="button sm primary" data-deck-save-layout>${icon("checkPlain")}保存布局</button>
        </div>
      </div>
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

const RENDER_SCALE = 1.6;
const HANDLE_RADIUS = 5.5;
const deckMetaCache = new Map();
const slideSceneCache = new Map();
const decodedImageCache = new Map();
const imageLoadInFlight = new Map();

function lruSet(map, key, value, maxEntries) {
  if (map.has(key)) map.delete(key);
  map.set(key, value);
  while (map.size > maxEntries) {
    const oldest = map.keys().next().value;
    map.delete(oldest);
  }
}

function loadCachedImage(src) {
  if (!src) return Promise.resolve(null);
  if (decodedImageCache.has(src)) {
    const img = decodedImageCache.get(src);
    decodedImageCache.delete(src);
    decodedImageCache.set(src, img);
    return Promise.resolve(img);
  }
  if (imageLoadInFlight.has(src)) return imageLoadInFlight.get(src);
  const promise = new Promise((resolve) => {
    const image = new Image();
    image.decoding = "async";
    image.src = src;
    image.decode().then(() => {
      lruSet(decodedImageCache, src, image, 80);
      resolve(image);
    }).catch(() => {
      if (image.complete && image.naturalWidth > 0) {
        lruSet(decodedImageCache, src, image, 80);
        resolve(image);
      } else {
        resolve(null);
      }
    }).finally(() => {
      imageLoadInFlight.delete(src);
    });
  });
  imageLoadInFlight.set(src, promise);
  return promise;
}

function getImmediateSceneImages(scene) {
  let hasPending = false;
  const images = (scene?.objects || []).map((shape) => {
    if (!shape.image) return null;
    const hit = decodedImageCache.get(shape.image);
    if (hit) return hit;
    hasPending = true;
    return null;
  });
  return { images, hasPending };
}

async function decodeSceneImages(scene) {
  return Promise.all((scene?.objects || []).map((shape) => loadCachedImage(shape.image)));
}

async function fetchSlideScene(fileId, number) {
  const key = `${fileId}:${number}`;
  if (slideSceneCache.has(key)) {
    const cached = slideSceneCache.get(key);
    slideSceneCache.delete(key);
    slideSceneCache.set(key, cached);
    return cached;
  }
  const data = await api("/files/" + encodeURIComponent(fileId) + "/slides/" + number);
  lruSet(slideSceneCache, key, data, 60);
  return data;
}

function prefetchAdjacentSlides(fileId, currentPage, total) {
  const targets = [];
  if (currentPage < total) targets.push(currentPage + 1);
  if (currentPage > 1) targets.push(currentPage - 1);
  for (const num of targets) {
    void fetchSlideScene(fileId, num)
      .then((nextScene) => {
        if (!nextScene?.objects) return;
        for (const obj of nextScene.objects) {
          if (obj.image) void loadCachedImage(obj.image);
        }
      })
      .catch(() => {});
  }
}

function effectiveRect(shape, overrides) {
  const ov = overrides?.get(shape.id);
  return ov || { left: shape.left, top: shape.top, width: shape.width, height: shape.height };
}

function resizeHandlesFor(box) {
  return [
    { key: "nw", x: box.left, y: box.top, cursor: "nwse-resize" },
    { key: "ne", x: box.left + box.width, y: box.top, cursor: "nesw-resize" },
    { key: "sw", x: box.left, y: box.top + box.height, cursor: "nesw-resize" },
    { key: "se", x: box.left + box.width, y: box.top + box.height, cursor: "nwse-resize" }
  ];
}

function hitHandle(box, x, y, hitPad = 10) {
  if (!box) return null;
  for (const h of resizeHandlesFor(box)) {
    if (Math.abs(x - h.x) <= hitPad && Math.abs(y - h.y) <= hitPad) return h;
  }
  return null;
}

function snapBoxToGuides(movingId, rawBox, scene, overrides, threshold = 6) {
  const guides = { x: [], y: [] };
  const xTargets = [0, scene.width / 2, scene.width, 51, scene.width - 51];
  const yTargets = [0, scene.height / 2, scene.height, 36, scene.height - 36];
  for (const item of scene.objects) {
    if (item.id === movingId || item.is_background) continue;
    const r = effectiveRect(item, overrides);
    xTargets.push(r.left, r.left + r.width / 2, r.left + r.width);
    yTargets.push(r.top, r.top + r.height / 2, r.top + r.height);
  }
  let left = rawBox.left;
  let top = rawBox.top;
  const w = rawBox.width;
  const h = rawBox.height;

  let bestDx = threshold + 1;
  let snappedX = null;
  for (const [edgeVal, offset] of [[left, 0], [left + w / 2, w / 2], [left + w, w]]) {
    for (const target of xTargets) {
      const d = Math.abs(edgeVal - target);
      if (d < bestDx) {
        bestDx = d;
        left = target - offset;
        snappedX = target;
      }
    }
  }
  if (snappedX !== null && bestDx <= threshold) guides.x.push(snappedX);

  let bestDy = threshold + 1;
  let snappedY = null;
  for (const [edgeVal, offset] of [[top, 0], [top + h / 2, h / 2], [top + h, h]]) {
    for (const target of yTargets) {
      const d = Math.abs(edgeVal - target);
      if (d < bestDy) {
        bestDy = d;
        top = target - offset;
        snappedY = target;
      }
    }
  }
  if (snappedY !== null && bestDy <= threshold) guides.y.push(snappedY);

  return { box: { left, top, width: w, height: h }, guides };
}

function paintInteractiveScene(canvas, scene, images, overrides, {
  selectedIds = [],
  hoveredId = null,
  regionBox = null,
  isDragging = false,
  mode = "select",
  guides = null,
  activeMoveId = null
} = {}) {
  if (!scene) return;
  const wPx = Math.round(scene.width * RENDER_SCALE);
  const hPx = Math.round(scene.height * RENDER_SCALE);
  if (canvas.width !== wPx || canvas.height !== hPx) {
    canvas.width = wPx;
    canvas.height = hPx;
  }
  const ctx = canvas.getContext("2d");
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.scale(RENDER_SCALE, RENDER_SCALE);

  ctx.fillStyle = scene.background;
  ctx.fillRect(0, 0, scene.width, scene.height);

  for (const [index, shape] of scene.objects.entries()) {
    const { left, top, width, height } = effectiveRect(shape, overrides);
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
    if (images?.[index]) {
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

  // 智能对齐参考线
  if (mode === "move" && guides) {
    ctx.save();
    ctx.strokeStyle = "#ec4899";
    ctx.lineWidth = 1;
    ctx.setLineDash([4, 3]);
    for (const gx of guides.x || []) {
      ctx.beginPath();
      ctx.moveTo(gx, 0);
      ctx.lineTo(gx, scene.height);
      ctx.stroke();
    }
    for (const gy of guides.y || []) {
      ctx.beginPath();
      ctx.moveTo(0, gy);
      ctx.lineTo(scene.width, gy);
      ctx.stroke();
    }
    ctx.restore();
  }

  // 拖动模式悬停提示框
  if (mode === "move" && hoveredId && hoveredId !== activeMoveId) {
    const hovered = scene.objects.find((o) => o.id === hoveredId);
    if (hovered) {
      const { left, top, width, height } = effectiveRect(hovered, overrides);
      ctx.save();
      ctx.strokeStyle = "rgba(37, 99, 235, 0.45)";
      ctx.lineWidth = 1.5;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(left, top, width, height);
      ctx.restore();
    }
  }

  const selectedSet = new Set(selectedIds.filter(Boolean));
  for (const shape of scene.objects) {
    if (!selectedSet.has(shape.id)) continue;
    const box = effectiveRect(shape, overrides);
    const { left, top, width, height } = box;
    ctx.save();
    if (mode === "move" && shape.id === activeMoveId) {
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 3.5;
      ctx.strokeRect(left, top, width, height);
      ctx.strokeStyle = "#2563eb";
      ctx.lineWidth = 2;
      ctx.strokeRect(left, top, width, height);
      for (const h of resizeHandlesFor(box)) {
        ctx.fillStyle = "#ffffff";
        ctx.strokeStyle = "#2563eb";
        ctx.lineWidth = 1.8;
        ctx.fillRect(h.x - HANDLE_RADIUS, h.y - HANDLE_RADIUS, HANDLE_RADIUS * 2, HANDLE_RADIUS * 2);
        ctx.strokeRect(h.x - HANDLE_RADIUS, h.y - HANDLE_RADIUS, HANDLE_RADIUS * 2, HANDLE_RADIUS * 2);
      }
      const xIn = (left / scene.width * SLIDE_W_IN).toFixed(2);
      const yIn = (top / scene.height * SLIDE_H_IN).toFixed(2);
      const wIn = (width / scene.width * SLIDE_W_IN).toFixed(2);
      const hIn = (height / scene.height * SLIDE_H_IN).toFixed(2);
      const tag = `X ${xIn}" · Y ${yIn}" · ${wIn}"×${hIn}"`;
      ctx.font = '600 10.5px "Microsoft YaHei", sans-serif';
      const tw = ctx.measureText(tag).width + 12;
      const th = 19;
      const tx = Math.max(4, Math.min(scene.width - tw - 4, left));
      const ty = top >= th + 8 ? top - th - 6 : top + 6;
      ctx.fillStyle = "#18181b";
      ctx.fillRect(tx, ty, tw, th);
      ctx.fillStyle = "#ffffff";
      ctx.textBaseline = "middle";
      ctx.textAlign = "left";
      ctx.fillText(tag, tx + 6, ty + th / 2);
    } else {
      ctx.strokeStyle = "rgba(255,255,255,0.9)";
      ctx.lineWidth = regionBox ? 2.5 : 4;
      ctx.strokeRect(left, top, width, height);
      ctx.strokeStyle = regionBox ? "rgba(37, 99, 235, 0.72)" : "#2563eb";
      ctx.lineWidth = regionBox ? 1.5 : 2.2;
      ctx.setLineDash([6, 4]);
      ctx.strokeRect(left, top, width, height);
      ctx.setLineDash([]);
    }
    ctx.restore();
  }

  if (mode === "select" && regionBox) {
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
  const modeBar = panel.querySelector("[data-deck-mode-bar]");
  const layoutActions = panel.querySelector("[data-deck-layout-actions]");
  const saveLayoutBtn = panel.querySelector("[data-deck-save-layout]");
  const resetLayoutBtn = panel.querySelector("[data-deck-reset-layout]");
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
  let decodedImages = [];
  let epoch = 0;
  let totalPages = 0;
  let interactionMode = "select"; // "select" | "move"
  let activeRegion = null; // { left, top, width, height } in scene coordinates
  let activeShapeIds = [];
  let activeLabel = "";
  let activeMoveId = null;
  let hoveredId = null;
  let layoutOverrides = new Map(); // shape_id -> { left, top, width, height }
  let activeGuides = null;
  let dragState = null; // { kind: "region"|"move"|"resize", ... }
  let savingLayout = false;

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

  function syncModeUi() {
    if (modeBar) {
      for (const btn of modeBar.querySelectorAll("[data-deck-mode]")) {
        const on = btn.dataset.deckMode === interactionMode;
        btn.classList.toggle("active", on);
        btn.setAttribute("aria-selected", on ? "true" : "false");
      }
    }
    if (layoutActions) {
      layoutActions.hidden = layoutOverrides.size === 0;
    }
    if (canvas) {
      canvas.classList.toggle("is-move-mode", interactionMode === "move");
      if (interactionMode === "select") canvas.style.cursor = "crosshair";
      else if (!dragState) canvas.style.cursor = "default";
    }
  }

  function setInteractionMode(nextMode) {
    interactionMode = nextMode === "move" ? "move" : "select";
    activeGuides = null;
    if (interactionMode === "move") {
      activeRegion = null;
      if (!activeMoveId && shape.value) activeMoveId = Number(shape.value);
      note.textContent = "拖动排版模式：直接拖动任意图片、卡片或文本框调整位置，拖动四角可缩放，支持方向键微调。";
    } else {
      hoveredId = null;
      note.textContent = activeRegion
        ? `已${activeLabel || "圈选区域"}，可在左侧直接输入修改要求（支持换图、改字、调排版）。`
        : "在画布上拖拽鼠标可自由圈选任意区域（含图片、卡片、文字），或单击任意元素直接让 AI 修改。";
    }
    syncModeUi();
    redrawCanvas(false);
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
    if (!scene) return;
    const selectedIds = interactionMode === "move"
      ? (activeMoveId ? [activeMoveId] : [])
      : (activeRegion ? activeShapeIds : (shape.value ? [Number(shape.value)] : []));
    paintInteractiveScene(canvas, scene, decodedImages, layoutOverrides, {
      selectedIds,
      hoveredId: interactionMode === "move" ? hoveredId : null,
      regionBox: interactionMode === "select" ? activeRegion : null,
      isDragging,
      mode: interactionMode,
      guides: activeGuides,
      activeMoveId
    });
    if (clearBoxBtn) clearBoxBtn.hidden = interactionMode !== "select" || (!activeRegion && !shape.value);
    if (layoutActions) layoutActions.hidden = layoutOverrides.size === 0;
  }

  function describeRegionHits(box) {
    if (!scene || !box) return { ids: [], label: "圈选区域" };
    const fgHits = [];
    let bgHit = null;
    for (const item of scene.objects) {
      const r = effectiveRect(item, layoutOverrides);
      if (!boxesIntersect(r, box)) continue;
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
    const hits = scene.objects.filter((item) => {
      const r = effectiveRect(item, layoutOverrides);
      return x >= r.left && x <= r.left + r.width && y >= r.top && y <= r.top + r.height;
    });
    if (!hits.length) return null;
    const foreground = hits.filter((item) => !item.is_background);
    if (foreground.length) {
      const score = (item) => {
        const r = effectiveRect(item, layoutOverrides);
        const kind = objectKind(item);
        const priority = kind === "text" ? 0 : kind === "image" ? 1 : 2;
        return priority * 1e9 + (r.width * r.height);
      };
      foreground.sort((a, b) => score(a) - score(b));
      return foreground[0];
    }
    return hits[hits.length - 1];
  }

  async function loadPptxPage(fileId, number) {
    const ticket = ++epoch;
    activeRegion = null;
    activeShapeIds = [];
    activeLabel = "";
    activeMoveId = null;
    hoveredId = null;
    activeGuides = null;
    layoutOverrides.clear();
    action.disabled = true;
    syncModeUi();
    syncPager(number);
    const cacheKey = `${fileId}:${number}`;
    const isCached = slideSceneCache.has(cacheKey);
    if (!isCached && loading) {
      loading.textContent = "正在渲染幻灯片…";
      loading.hidden = false;
    }
    const next = await fetchSlideScene(fileId, number);
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
        activeMoveId = Number(selected.shape_id);
        activeLabel = selected.label || "";
      }
    }
    // 增量秒绘：先立即画出所有矢量排版、文字与已缓存图片，未解码图片后台并行补绘，告别白屏等待
    const immediate = getImmediateSceneImages(next);
    decodedImages = immediate.images;
    redrawCanvas(false);
    canvas.hidden = false;
    if (!isCached) {
      canvas.classList.remove("fading");
      void canvas.offsetWidth;
      canvas.classList.add("fading");
    }
    if (loading) loading.hidden = true;
    canvas.setAttribute("aria-label", `第 ${number} 页预览（支持框选修改或拖动排版）`);
    shape.disabled = false;
    action.disabled = !document.querySelector("[data-form=conversation]");
    syncModeUi();
    note.textContent = interactionMode === "move"
      ? "拖动排版模式：直接拖动任意图片、卡片或文本框调整位置，拖动四角可缩放，支持方向键微调。"
      : (activeRegion
        ? `已${activeLabel || "圈选区域"}，可在左侧直接输入修改要求（支持换图、改字、调排版）。`
        : "在画布上拖拽鼠标可自由圈选任意区域（含图片、卡片、文字），或单击任意元素直接让 AI 修改。");
    remember(currentWorkspace, { deck: { fileId, page: number } });

    if (immediate.hasPending) {
      void decodeSceneImages(next).then((fullImages) => {
        if (!panel.isConnected || ticket !== epoch || scene !== next) return;
        decodedImages = fullImages;
        redrawCanvas(false);
      });
    }
    prefetchAdjacentSlides(fileId, number, totalPages);
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
    let result = deckMetaCache.get(fileId);
    if (!result) {
      if (loading) { loading.textContent = "正在加载演示文稿…"; loading.hidden = false; }
      result = await api("/files/" + encodeURIComponent(fileId) + "/slides");
      lruSet(deckMetaCache, fileId, result, 16);
      if (result?.first_page && result.first_page.page === 1) {
        lruSet(slideSceneCache, `${fileId}:1`, result.first_page, 60);
      }
    }
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
    activeMoveId = sid;
    const obj = sid && scene ? scene.objects.find((item) => item.id === sid) : null;
    activeLabel = obj ? describeSingleObject(obj) : "";
    redrawCanvas(false);
    if (interactionMode === "select") {
      if (sid) applySelection(false);
      else clearSelection();
    }
  });
  modeBar?.addEventListener("click", (event) => {
    const btn = event.target.closest("[data-deck-mode]");
    if (!btn) return;
    setInteractionMode(btn.dataset.deckMode);
  });
  resetLayoutBtn?.addEventListener("click", () => {
    layoutOverrides.clear();
    activeGuides = null;
    redrawCanvas(false);
    toastMessage("已还原当前页手动拖动改动", "info");
  });
  saveLayoutBtn?.addEventListener("click", async () => {
    if (savingLayout || !scene || !layoutOverrides.size) return;
    const currentFileId = file.value;
    const currentPageNum = Number(page.value) || 1;
    if (!currentFileId) return;
    savingLayout = true;
    saveLayoutBtn.disabled = true;
    saveLayoutBtn.textContent = "正在保存…";
    try {
      const changes = [];
      for (const [shapeId, box] of layoutOverrides.entries()) {
        changes.push({
          shape_id: Number(shapeId),
          x: Number((box.left / scene.width * SLIDE_W_IN).toFixed(3)),
          y: Number((box.top / scene.height * SLIDE_H_IN).toFixed(3)),
          w: Number(Math.max(0.08, box.width / scene.width * SLIDE_W_IN).toFixed(3)),
          h: Number(Math.max(0.08, box.height / scene.height * SLIDE_H_IN).toFixed(3))
        });
      }
      const res = await post(`/files/${encodeURIComponent(currentFileId)}/slides/${currentPageNum}/layout`, { changes });
      const newFile = res?.file;
      if (newFile?.id) {
        // 同步替换当前会话内存中的对应 turn.file，让画布下拉与下载按钮立刻指向新版 PPT
        const turns = currentWorkspace?.session?.turns || [];
        for (let i = turns.length - 1; i >= 0; i -= 1) {
          if (turns[i].file?.id === currentFileId) {
            turns[i].file = newFile;
            break;
          }
        }
        layoutOverrides.clear();
        activeGuides = null;
        await activeController?.syncArtifacts(currentWorkspace, { selectId: newFile.id, preferredPage: currentPageNum, flash: true });
        toastMessage(`已保存第 ${currentPageNum} 页布局（已调整 ${res.updated || changes.length} 个组件）`, "success");
      }
    } catch (error) {
      toastMessage(error?.message || "布局保存失败", "error");
    } finally {
      savingLayout = false;
      if (saveLayoutBtn.isConnected) {
        saveLayoutBtn.disabled = false;
        saveLayoutBtn.innerHTML = `${icon("checkPlain")}保存布局`;
      }
    }
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
    if (!scene || event.button !== 0) return;
    event.preventDefault();
    const pt = scenePointFromEvent(event);
    try { canvas.setPointerCapture(event.pointerId); } catch {}

    if (interactionMode === "move") {
      // 1. 先检查是否点中当前选中组件的四角缩放手柄
      if (activeMoveId) {
        const curObj = scene.objects.find((o) => o.id === activeMoveId);
        if (curObj) {
          const curBox = effectiveRect(curObj, layoutOverrides);
          const handle = hitHandle(curBox, pt.x, pt.y);
          if (handle) {
            dragState = { kind: "resize", id: curObj.id, handle: handle.key, startPt: pt, startBox: { ...curBox } };
            canvas.style.cursor = handle.cursor;
            return;
          }
        }
      }
      // 2. 命中组件即可直接拖动移动
      const hit = pickObjectAtPoint(pt.x, pt.y);
      if (hit) {
        activeMoveId = hit.id;
        shape.value = String(hit.id);
        const startBox = { ...effectiveRect(hit, layoutOverrides) };
        dragState = { kind: "move", id: hit.id, startPt: pt, startBox, moved: false };
        canvas.style.cursor = "grabbing";
        note.textContent = `已选中「${describeSingleObject(hit)}」，拖动移动位置，拖四角缩放大小。`;
        redrawCanvas(false);
      } else {
        activeMoveId = null;
        shape.value = "";
        dragState = null;
        redrawCanvas(false);
      }
      return;
    }

    // 框选给 AI 模式
    dragState = { kind: "region", startPt: pt, dragging: false };
  });

  canvas.addEventListener("pointermove", (event) => {
    if (!scene) return;
    const pt = scenePointFromEvent(event);

    if (!dragState) {
      if (interactionMode === "move") {
        if (activeMoveId) {
          const curObj = scene.objects.find((o) => o.id === activeMoveId);
          if (curObj) {
            const handle = hitHandle(effectiveRect(curObj, layoutOverrides), pt.x, pt.y);
            if (handle) {
              canvas.style.cursor = handle.cursor;
              return;
            }
          }
        }
        const hit = pickObjectAtPoint(pt.x, pt.y);
        const nextHover = hit ? hit.id : null;
        canvas.style.cursor = hit ? "move" : "default";
        if (nextHover !== hoveredId) {
          hoveredId = nextHover;
          redrawCanvas(false);
        }
      }
      return;
    }

    if (dragState.kind === "move") {
      const dx = pt.x - dragState.startPt.x;
      const dy = pt.y - dragState.startPt.y;
      if (!dragState.moved && Math.hypot(dx, dy) < 2) return;
      dragState.moved = true;
      const rawBox = {
        left: Math.max(-40, Math.min(scene.width - 10, dragState.startBox.left + dx)),
        top: Math.max(-40, Math.min(scene.height - 10, dragState.startBox.top + dy)),
        width: dragState.startBox.width,
        height: dragState.startBox.height
      };
      const snapped = snapBoxToGuides(dragState.id, rawBox, scene, layoutOverrides);
      layoutOverrides.set(dragState.id, snapped.box);
      activeGuides = snapped.guides;
      redrawCanvas(true);
      return;
    }

    if (dragState.kind === "resize") {
      const dx = pt.x - dragState.startPt.x;
      const dy = pt.y - dragState.startPt.y;
      const sb = dragState.startBox;
      const minW = 20;
      const minH = 14;
      let left = sb.left;
      let top = sb.top;
      let right = sb.left + sb.width;
      let bottom = sb.top + sb.height;
      if (dragState.handle.includes("w")) left = Math.min(right - minW, Math.max(0, sb.left + dx));
      if (dragState.handle.includes("e")) right = Math.max(left + minW, Math.min(scene.width, sb.left + sb.width + dx));
      if (dragState.handle.includes("n")) top = Math.min(bottom - minH, Math.max(0, sb.top + dy));
      if (dragState.handle.includes("s")) bottom = Math.max(top + minH, Math.min(scene.height, sb.top + sb.height + dy));
      layoutOverrides.set(dragState.id, { left, top, width: right - left, height: bottom - top });
      activeGuides = null;
      redrawCanvas(true);
      return;
    }

    if (dragState.kind === "region") {
      const dist = Math.hypot(pt.x - dragState.startPt.x, pt.y - dragState.startPt.y);
      if (!dragState.dragging && dist < 8) return;
      dragState.dragging = true;
      const box = normalizeBox(dragState.startPt, pt, scene.width, scene.height);
      activeRegion = box;
      const hitInfo = describeRegionHits(box);
      activeShapeIds = hitInfo.ids;
      activeLabel = hitInfo.label;
      shape.value = "";
      redrawCanvas(true);
    }
  });

  function finishPointer(event, cancelled = false) {
    if (!dragState || !scene) return;
    const stateSnap = dragState;
    dragState = null;
    activeGuides = null;
    try { canvas.releasePointerCapture(event.pointerId); } catch {}

    if (stateSnap.kind === "move" || stateSnap.kind === "resize") {
      canvas.style.cursor = "move";
      redrawCanvas(false);
      if (layoutOverrides.size > 0) {
        note.textContent = `已调整 ${layoutOverrides.size} 个组件位置/尺寸，点击上方「保存布局」写入 PPTX。`;
      }
      return;
    }

    if (cancelled) {
      redrawCanvas(false);
      return;
    }
    const end = scenePointFromEvent(event);
    if (stateSnap.dragging && Math.abs(end.x - stateSnap.startPt.x) >= 10 && Math.abs(end.y - stateSnap.startPt.y) >= 10) {
      const box = normalizeBox(stateSnap.startPt, end, scene.width, scene.height);
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
    activeRegion = null;
    const hit = pickObjectAtPoint(end.x, end.y);
    if (hit) {
      shape.value = String(hit.id);
      activeShapeIds = [hit.id];
      activeMoveId = hit.id;
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
  canvas.addEventListener("pointerleave", () => {
    if (!dragState && hoveredId !== null) {
      hoveredId = null;
      redrawCanvas(false);
    }
  });

  // 拖动排版模式下支持方向键精准微调组件位置
  const onKeydown = (event) => {
    if (!panel.isConnected) {
      document.removeEventListener("keydown", onKeydown);
      return;
    }
    if (interactionMode !== "move" || !activeMoveId || !scene) return;
    if (event.target?.matches?.("input, textarea, select, [contenteditable]")) return;
    const step = event.shiftKey ? 12 : 2;
    let dx = 0;
    let dy = 0;
    if (event.key === "ArrowLeft") dx = -step;
    else if (event.key === "ArrowRight") dx = step;
    else if (event.key === "ArrowUp") dy = -step;
    else if (event.key === "ArrowDown") dy = step;
    else return;
    event.preventDefault();
    const curObj = scene.objects.find((o) => o.id === activeMoveId);
    if (!curObj) return;
    const box = effectiveRect(curObj, layoutOverrides);
    layoutOverrides.set(activeMoveId, {
      left: Math.max(-40, Math.min(scene.width - 10, box.left + dx)),
      top: Math.max(-40, Math.min(scene.height - 10, box.top + dy)),
      width: box.width,
      height: box.height
    });
    redrawCanvas(false);
  };
  document.addEventListener("keydown", onKeydown);

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
