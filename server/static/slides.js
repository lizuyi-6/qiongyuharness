import { api, button, e, icon, readDraft, saveDraft, state, toastMessage } from "./ui.js";

export function deckFiles(session) {
  return [...new Map((session?.turns || []).filter((turn) => turn.file?.extension === "pptx")
    .map((turn) => [turn.file.id, turn.file])).values()].reverse();
}

export function selectionChip(selection) {
  return selection ? `<span class="attached-file">第 ${selection.slide} 页${selection.shape_id ? " · 选中文本框" : " · 整页文字"}
    <button type="button" data-act="slides-clear" aria-label="清除选区">×</button></span>` : "";
}

export function deckPanel(workspace) {
  const decks = deckFiles(workspace.session);
  if (!decks.length) return "";
  const draft = readDraft(workspace.draftKey) || {};
  const version = decks.find((file) => file.id === draft.deck?.fileId) || decks[0];
  return `<section class="deck-studio" data-deck-preview aria-label="PPT 预览画布">
    <div class="deck-toolbar">
      <div class="deck-file-wrap">
        <select class="deck-select" data-deck-file aria-label="选择演示文稿版本">${decks.map((file, index) =>
          `<option value="${e(file.id)}" ${file.id === version.id ? "selected" : ""}>${e(file.name)}${decks.length > 1 ? (index === 0 ? "（最新）" : `（旧版 ${index}）`) : ""}</option>`).join("")}</select>
      </div>
      <a class="button sm" href="/api/v1/files/${encodeURIComponent(version.id)}/download" data-deck-download download>${icon("download")}下载</a>
    </div>

    <div class="deck-stage">
      <div class="deck-canvas-frame">
        <canvas data-deck-canvas role="img" aria-label="幻灯片预览图" hidden></canvas>
        <p class="deck-loading hint" data-deck-loading role="status">正在加载幻灯片…</p>
      </div>
    </div>

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
        ${button("引用选区修改", "slides-target", { cls: "primary sm", disabled: true })}
      </div>
      <p class="hint deck-note" data-deck-note role="status">点击画布中任意文本框可直接定位选区。</p>
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
    if (shape.fill) {
      ctx.fillStyle = shape.fill;
      ctx.fillRect(left, top, width, height);
    }
    if (images[index]) ctx.drawImage(images[index], left, top, width, height);
    if (shape.text) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(left, top, width, height);
      ctx.clip();
      ctx.font = `${shape.bold ? "700 " : ""}${Math.max(7, shape.size)}px "Microsoft YaHei","Noto Sans CJK SC",sans-serif`;
      ctx.fillStyle = shape.color;
      ctx.textBaseline = "top";
      textLines(ctx, shape.text, width - 3).slice(0, 100)
        .forEach((line, lineIndex) => ctx.fillText(line, left + 1, top + 2 + lineIndex * shape.size * 1.21));
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

export function selectDeckFile(fileId) {
  const panel = document.querySelector("[data-deck-preview]");
  const file = panel?.querySelector("[data-deck-file]");
  if (!file || !fileId) return;
  if ([...file.options].some((opt) => opt.value === fileId) && file.value !== fileId) {
    file.value = fileId;
    file.dispatchEvent(new Event("change"));
  }
}

export function mountDeck() {
  const panel = document.querySelector("[data-deck-preview]");
  const workspace = state.chatWorkspace;
  if (!panel || !workspace) return;
  const file = panel.querySelector("[data-deck-file]");
  const page = panel.querySelector("[data-deck-page]");
  const prevBtn = panel.querySelector("[data-deck-prev]");
  const nextBtn = panel.querySelector("[data-deck-next]");
  const counter = panel.querySelector("[data-deck-counter]");
  const shape = panel.querySelector("[data-deck-shape]");
  const canvas = panel.querySelector("canvas");
  const loading = panel.querySelector("[data-deck-loading]");
  const note = panel.querySelector("[data-deck-note]");
  const action = panel.querySelector("[data-act=slides-target]");
  const saved = readDraft(workspace.draftKey) || {};
  let scene, epoch = 0, totalPages = 0;
  const root = () => "/files/" + encodeURIComponent(file.value) + "/slides";

  function syncPager(current) {
    if (counter) {
      counter.textContent = totalPages ? `${String(current).padStart(2, "0")} / ${String(totalPages).padStart(2, "0")}` : "— / —";
    }
    if (prevBtn) prevBtn.disabled = !totalPages || current <= 1;
    if (nextBtn) nextBtn.disabled = !totalPages || current >= totalPages;
  }

  async function loadPage(number) {
    const ticket = ++epoch;
    const id = file.value;
    scene = null;
    action.disabled = true;
    syncPager(number);
    if (loading) { loading.textContent = "正在读取幻灯片…"; loading.hidden = false; }
    const next = await api(root() + "/" + number);
    if (!panel.isConnected || ticket !== epoch) return;
    scene = next;
    const objects = scene.objects.filter((item) => item.editable);
    rebuildChoices(shape, [{ text: "整页的可编辑文字", value: "" }, ...objects.map((item) => ({
      text: item.text.slice(0, 31), value: String(item.id)
    }))]);
    const selected = (readDraft(workspace.draftKey) || {}).selection;
    if (selected?.file_id === id && selected.slide === number && selected.shape_id) {
      shape.value = String(selected.shape_id);
    }
    await art(canvas, next, Number(shape.value));
    if (!panel.isConnected || ticket !== epoch) return;
    canvas.hidden = false;
    if (loading) loading.hidden = true;
    canvas.setAttribute("aria-label", `第 ${number} 页预览`);
    shape.disabled = !objects.length;
    action.disabled = !(objects.length && document.querySelector("[data-form=conversation]"));
    note.textContent = objects.length ? "可直接点击画布中的文本框，或下拉选择修改范围。" : "当前页没有可编辑的单段文本框。";
    remember(workspace, { deck: { fileId: id, page: number } });
  }

  async function loadFile(preferred = 1) {
    const ticket = ++epoch;
    scene = null;
    totalPages = 0;
    page.disabled = true;
    shape.disabled = true;
    action.disabled = true;
    canvas.hidden = true;
    syncPager(0);
    if (loading) { loading.textContent = "正在加载预览…"; loading.hidden = false; }
    const result = await api(root());
    if (!panel.isConnected || ticket !== epoch) return;
    totalPages = result.pages.length;
    rebuildChoices(page, result.pages.map((slide) => ({ text: `${slide.page}. ${slide.title}`, value: String(slide.page) })));
    const targetPage = Math.max(1, Math.min(preferred, totalPages || 1));
    page.value = String(targetPage);
    page.disabled = totalPages <= 1;
    panel.querySelector("[data-deck-download]").href = "/api/v1/files/" + encodeURIComponent(result.id) + "/download";
    await loadPage(targetPage);
  }

  function report(error) {
    if (panel.isConnected) {
      if (loading) loading.hidden = true;
      note.textContent = error.message;
    }
  }

  file.addEventListener("change", () => { clearSelection(); void loadFile().catch(report); });
  page.addEventListener("change", () => { clearSelection(); void loadPage(Number(page.value)).catch(report); });
  prevBtn?.addEventListener("click", () => {
    const cur = Number(page.value);
    if (cur > 1) { page.value = String(cur - 1); clearSelection(); void loadPage(cur - 1).catch(report); }
  });
  nextBtn?.addEventListener("click", () => {
    const cur = Number(page.value);
    if (cur < totalPages) { page.value = String(cur + 1); clearSelection(); void loadPage(cur + 1).catch(report); }
  });
  shape.addEventListener("change", () => { clearSelection(); if (scene) void art(canvas, scene, Number(shape.value)); });
  canvas.addEventListener("click", (event) => {
    if (!scene || shape.disabled) return;
    const bounds = canvas.getBoundingClientRect();
    const x = (event.clientX - bounds.left) * scene.width / bounds.width;
    const y = (event.clientY - bounds.top) * scene.height / bounds.height;
    const hit = [...scene.objects].reverse().find((item) => item.editable && x >= item.left && x <= item.left + item.width && y >= item.top && y <= item.top + item.height);
    if (hit) { clearSelection(); shape.value = String(hit.id); void art(canvas, scene, hit.id); }
  });
  action.addEventListener("click", () => {
    const form = document.querySelector("[data-form=conversation]");
    if (!form || form.dataset.busy || !scene) return;
    const selection = { file_id: file.value, slide: Number(page.value), shape_id: shape.value ? Number(shape.value) : null };
    remember(workspace, { selection });
    form.querySelector("[data-chat-selection]").innerHTML = selectionChip(selection);
    form.elements.message.focus();
    toastMessage("已选中要修改的幻灯片范围", "success");
  });
  void loadFile(saved.deck?.page || 1).catch(report);
}
