import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const resolve = createRequire(process.env.QO_DSH_EXECUTABLE || import.meta.url);
const { defineTool } = await import(pathToFileURL(resolve.resolve("@deepseek-ai/dsh-tools")).href);
const ALLOWED_TOOLS = new Set(["office_context_read", "office_image_generate", "office_ppt_check", "web_search", "web_fetch"]);
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
const CJK = /[\u2E80-\u9FFF\uF900-\uFAFF\uFE30-\uFE4F\u3000-\u303F]/;
const GENERIC_TITLES = /^(公司概况|项目概况|核心内容|主要内容|应用与案例|应用场景|总结与启示|总结|概述|介绍|发展历程|产品介绍|背景介绍|结论)$/;

export const name = "qiongyu-office-ppt-check";
export const inject = ["tools"];

function textEm(text) {
  let total = 0;
  for (const ch of String(text || "")) total += CJK.test(ch) ? 1.0 : 0.56;
  return total;
}

function isDarkHex(value) {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(String(value || "").trim());
  if (!m) return false;
  const hex = m[1];
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) < 142;
}

function overlapArea(a, b) {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.w, b.x + b.w);
  const y2 = Math.min(a.y + a.h, b.y + b.h);
  if (x2 <= x1 || y2 <= y1) return 0;
  return (x2 - x1) * (y2 - y1);
}

function boxContainsCenter(outer, inner) {
  const cx = inner.x + inner.w / 2;
  const cy = inner.y + inner.h / 2;
  return cx >= outer.x - 0.05 && cx <= outer.x + outer.w + 0.05 && cy >= outer.y - 0.05 && cy <= outer.y + outer.h + 0.05;
}

function checkDeck(args) {
  let deck = args?.deck;
  if (typeof deck === "string") {
    try { deck = JSON.parse(deck); } catch {
      return "校验失败：deck 不是合法 JSON 对象";
    }
  }
  if (!deck || typeof deck !== "object" || !Array.isArray(deck.slides) || !deck.slides.length) {
    return "校验失败：缺少有效的 deck.slides 数组";
  }
  const issues = [];
  const warnings = [];
  const totalSlides = deck.slides.length;
  if (totalSlides > 30) issues.push(`总页数 ${totalSlides} 超过上限 30 页`);

  const paletteBg = deck?.design?.palette?.bg || "#FFFFFF";
  let darkCount = 0;
  let lightCount = 0;
  let chartOrTableCount = 0;
  let imageCount = 0;
  let prevSig = "";

  deck.slides.forEach((slide, idx) => {
    const page = idx + 1;
    const isCover = idx === 0;
    const bg = slide?.bg || paletteBg;
    if (isDarkHex(bg)) darkCount += 1;
    else lightCount += 1;
    if (slide?.bg_image) imageCount += 1;

    const elements = Array.isArray(slide?.elements) ? slide.elements : [];
    if (!elements.length) {
      issues.push(`第 ${page} 页没有任何 elements`);
      return;
    }

    const slideTitle = String(slide?.title || "").trim();
    if (!isCover && GENERIC_TITLES.test(slideTitle)) {
      warnings.push(`第 ${page} 页标题「${slideTitle}」过于空洞，请改为一句话核心结论（Action Title）`);
    }

    const textBoxes = [];
    const cards = [];
    let hasHeaderText = false;
    let hasChart = false;
    let hasTable = false;
    let hasImage = Boolean(slide?.bg_image);

    elements.forEach((el, eIdx) => {
      if (!el || typeof el !== "object") return;
      const type = String(el.type || "").toLowerCase();
      const x = Number(el.x);
      const y = Number(el.y);
      const w = Number(el.w);
      const h = Number(el.h);
      if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0) {
        issues.push(`第 ${page} 页元素 #${eIdx + 1}(${type}) 坐标或尺寸无效`);
        return;
      }
      if (type === "text") {
        if (y < 1.35) hasHeaderText = true;
        if (x < 0.25 || x + w > SLIDE_W - 0.25 || y < 0.2 || y + h > SLIDE_H - 0.15) {
          issues.push(`第 ${page} 页文本框 #${eIdx + 1}（"${String(el.text || "").slice(0, 12)}"）贴边或超出画布安全区`);
        }
        const size = Number(el.size) || 16;
        const isFooterOrTag = y >= 6.8 || (y < 0.72 && size >= 9.5 && String(el.text || "").length <= 24);
        if (!isFooterOrTag && size < 12.5) {
          issues.push(`第 ${page} 页文本框 #${eIdx + 1}（"${String(el.text || "").slice(0, 12)}"）字号仅 ${size}pt 过小（蚂蚁字），正文建议 13.5–15pt`);
        }
        const lineSpacing = Number(el.line_spacing) || 1.28;
        const linesRaw = String(el.text || "").split("\n");
        const perLine = Math.max(1, (w * 72 / size) * 0.92);
        const totalLines = linesRaw.reduce((sum, part) => sum + Math.max(1, Math.ceil(textEm(part) / perLine)), 0);
        const needH = totalLines * size * Math.max(1.15, lineSpacing) / 72;
        if (needH > h * 1.22) {
          warnings.push(`第 ${page} 页文本框 #${eIdx + 1}（"${String(el.text || "").slice(0, 12)}"）内容偏长，估算需高度 ${needH.toFixed(2)}" > 给定 ${h.toFixed(2)}"，建议精简文字或加大 h`);
        }
        textBoxes.push({ idx: eIdx + 1, x, y, w, h, text: String(el.text || ""), snippet: String(el.text || "").slice(0, 10) });
      } else if (type === "round_rect" || type === "rect") {
        if (w >= 2.6 && h >= 2.2 && w < 12.8 && h < 6.8) {
          cards.push({ idx: eIdx + 1, x, y, w, h, area: w * h });
        }
      } else if (type === "image") {
        hasImage = true;
        imageCount += 1;
      } else if (type === "table") {
        hasTable = true;
        chartOrTableCount += 1;
      } else if (type === "chart" || type === "hbar" || type === "vbar" || type === "kpi_row") {
        hasChart = true;
        chartOrTableCount += 1;
      }
    });

    if (!isCover && !hasHeaderText && !slide?.bg_image) {
      issues.push(`第 ${page} 页顶部缺少主标题（y < 1.35 处无文字），请补充 Eyebrow 分类标签与观点型大标题`);
    }

    // 文本框重叠检查
    for (let i = 0; i < textBoxes.length; i += 1) {
      for (let j = i + 1; j < textBoxes.length; j += 1) {
        const area = overlapArea(textBoxes[i], textBoxes[j]);
        const minArea = Math.min(textBoxes[i].w * textBoxes[i].h, textBoxes[j].w * textBoxes[j].h);
        if (area > 0.08 && area / Math.max(0.01, minArea) > 0.12) {
          issues.push(`第 ${page} 页文本框 #${textBoxes[i].idx}("${textBoxes[i].snippet}") 与 #${textBoxes[j].idx}("${textBoxes[j].snippet}") 坐标重叠`);
        }
      }
    }

    // 傻大空卡片检查：大面积卡片内只有寥寥几十字
    for (const card of cards) {
      if (card.area >= 12.0) {
        const innerChars = textBoxes
          .filter((tb) => boxContainsCenter(card, tb))
          .reduce((sum, tb) => sum + tb.text.trim().length, 0);
        if (innerChars > 0 && innerChars < 42) {
          issues.push(`第 ${page} 页卡片 #${card.idx}（${card.w.toFixed(1)}×${card.h.toFixed(1)}"）面积过大但内容仅 ${innerChars} 字，显得空洞廉价！请收紧卡片高度或补充关键指标与具体要点`);
        }
      }
    }

    // 连续页撞版检查
    const sig = `${hasChart ? "C" : ""}${hasTable ? "T" : ""}${hasImage ? "I" : ""}-cards${cards.length}`;
    if (!isCover && idx > 1 && cards.length >= 2 && sig === prevSig && !hasChart && !hasTable && !hasImage) {
      issues.push(`第 ${page} 页与第 ${page - 1} 页版式雷同（连续都是 ${cards.length} 张纯文字卡片），请改为图表 chart、表格 table、时间轴或图文混排拉开视觉节奏`);
    }
    prevSig = sig;
  });

  if (totalSlides >= 5 && lightCount === 0) {
    issues.push(`整套 ${totalSlides} 页全为深黑底色，缺乏明暗对比与杂志高级感！请将大部分正文页改为高级浅色底（如 #F8FAFC / #FAF9F5）配白卡片，仅保留封面与 1–2 页重点页为深色反转`);
  }
  if (totalSlides >= 5 && chartOrTableCount === 0) {
    warnings.push(`整套 ${totalSlides} 页没有任何原生数据图表（chart）或对比表格（table），强烈建议把至少 1–2 组关键数据升级为 chart(hbar/vbar/kpi_row) 或 table 提升专业说服力`);
  }

  if (!issues.length && !warnings.length) {
    return `版式与审美自检全部通过：共 ${totalSlides} 页（浅色页 ${lightCount} / 深色页 ${darkCount} / 图表与表格 ${chartOrTableCount} / 配图 ${imageCount}），可直接输出最终 JSON。`;
  }
  const lines = [`版式与审美自检结果（共 ${totalSlides} 页，浅色 ${lightCount} / 深色 ${darkCount} / 图表与表格 ${chartOrTableCount}）：`];
  if (issues.length) lines.push("必须修复的问题（请直接按以上坐标/字号修正后立即输出最终完整 JSON，无需再次重复调用 office_ppt_check）：", ...issues.slice(0, 16).map((s) => `- ${s}`));
  if (warnings.length) lines.push("建议优化项（可顺手微调后直接输出最终 JSON，无需再次调用校验）：", ...warnings.slice(0, 10).map((s) => `- ${s}`));
  return lines.join("\n");
}

export function apply(ctx) {
  ctx.tools.guard((execution) => ALLOWED_TOOLS.has(execution.name) ? undefined : "本轮只允许调用办公资料读取、联网搜索、生图与 PPT 排版校验工具");
  ctx.tools.register(defineTool({
    name: "office_ppt_check",
    description: "在输出最终 PPT 自由画布 JSON 前必须调用，自动检测每一页的坐标越界、文本框重叠、蚂蚁字、空洞大卡片、全套死黑单调底色及连续撞版，返回精准修改建议。",
    parameters: {
      deck: { type: "json", description: "必填。待校验的 deck 自由画布对象（含 title、design、slides）。" }
    },
    timeoutMs: 15_000,
    output: { schema: { type: "string" }, render: (_args, value) => [{ type: "text", text: value }] },
    execute: (args) => checkDeck(args)
  }));
}
