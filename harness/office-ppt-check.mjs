import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const resolve = createRequire(process.env.QO_DSH_EXECUTABLE || import.meta.url);
const { defineTool } = await import(pathToFileURL(resolve.resolve("@deepseek-ai/dsh-tools")).href);
const ALLOWED_TOOLS = new Set(["office_context_read", "office_image_generate", "office_ppt_check", "web_search", "web_fetch"]);
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
const CJK = /[\u2E80-\u9FFF\uF900-\uFAFF\uFE30-\uFE4F\u3000-\u303F]/;

export const name = "qiongyu-office-ppt-check";
export const inject = ["tools"];

function textEm(text) {
  let total = 0;
  for (const ch of String(text || "")) total += CJK.test(ch) ? 1.0 : 0.56;
  return total;
}

function overlapArea(a, b) {
  const x1 = Math.max(a.x, b.x);
  const y1 = Math.max(a.y, b.y);
  const x2 = Math.min(a.x + a.w, b.x + b.w);
  const y2 = Math.min(a.y + a.h, b.y + b.h);
  if (x2 <= x1 || y2 <= y1) return 0;
  return (x2 - x1) * (y2 - y1);
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

  deck.slides.forEach((slide, idx) => {
    const page = idx + 1;
    const elements = Array.isArray(slide?.elements) ? slide.elements : [];
    if (!elements.length) {
      issues.push(`第 ${page} 页没有任何 elements`);
      return;
    }
    const textBoxes = [];
    let hasVisual = Boolean(slide?.bg_image);
    elements.forEach((el, eIdx) => {
      if (!el || typeof el !== "object") return;
      const x = Number(el.x);
      const y = Number(el.y);
      const w = Number(el.w);
      const h = Number(el.h);
      if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0) {
        issues.push(`第 ${page} 页元素 #${eIdx + 1}(${el.type}) 坐标或尺寸无效`);
        return;
      }
      if (el.type === "text") {
        if (x < 0.25 || x + w > SLIDE_W - 0.25 || y < 0.2 || y + h > SLIDE_H - 0.2) {
          issues.push(`第 ${page} 页文本框 #${eIdx + 1}（"${String(el.text || "").slice(0, 12)}"）贴边或超出画布安全区`);
        }
        const size = Number(el.size) || 16;
        const lineSpacing = Number(el.line_spacing) || 1.25;
        const linesRaw = String(el.text || "").split("\n");
        const perLine = Math.max(1, (w * 72 / size) * 0.92);
        const totalLines = linesRaw.reduce((sum, part) => sum + Math.max(1, Math.ceil(textEm(part) / perLine)), 0);
        const needH = totalLines * size * Math.max(1.15, lineSpacing) / 72;
        if (needH > h * 1.25) {
          warnings.push(`第 ${page} 页文本框 #${eIdx + 1}（"${String(el.text || "").slice(0, 12)}"）内容偏长，估算需高度 ${needH.toFixed(2)}" > 给定 ${h.toFixed(2)}"，建议精简文字或加大 h`);
        }
        textBoxes.push({ idx: eIdx + 1, x, y, w, h, snippet: String(el.text || "").slice(0, 10) });
      } else if (["rect", "round_rect", "ellipse", "image", "table"].includes(el.type)) {
        hasVisual = true;
      }
    });

    for (let i = 0; i < textBoxes.length; i += 1) {
      for (let j = i + 1; j < textBoxes.length; j += 1) {
        const area = overlapArea(textBoxes[i], textBoxes[j]);
        const minArea = Math.min(textBoxes[i].w * textBoxes[i].h, textBoxes[j].w * textBoxes[j].h);
        if (area > 0.08 && area / Math.max(0.01, minArea) > 0.12) {
          issues.push(`第 ${page} 页文本框 #${textBoxes[i].idx}("${textBoxes[i].snippet}") 与 #${textBoxes[j].idx}("${textBoxes[j].snippet}") 坐标重叠`);
        }
      }
    }
    if (!hasVisual && textBoxes.length <= 2) {
      warnings.push(`第 ${page} 页视觉元素偏少，可补充色块、卡片、线条或配图提升层次`);
    }
  });

  if (!issues.length && !warnings.length) {
    return `版式自检通过：共 ${totalSlides} 页，未发现越界、文本框重叠或明显溢出问题，可以直接输出最终 JSON。`;
  }
  const lines = [`版式自检结果（共 ${totalSlides} 页）：`];
  if (issues.length) lines.push("必须修复的问题：", ...issues.slice(0, 16).map((s) => `- ${s}`));
  if (warnings.length) lines.push("建议优化项：", ...warnings.slice(0, 10).map((s) => `- ${s}`));
  return lines.join("\n");
}

export function apply(ctx) {
  ctx.tools.guard((execution) => ALLOWED_TOOLS.has(execution.name) ? undefined : "本轮只允许调用办公资料读取、生图与 PPT 排版校验工具");
  ctx.tools.register(defineTool({
    name: "office_ppt_check",
    description: "在输出最终 PPT 自由画布 JSON 前调用，自动检测每一页是否存在坐标越界、文本框重叠或字数溢出，返回精准修改建议。",
    parameters: {
      deck: { type: "object", required: true, description: "待校验的 deck 自由画布对象（含 title、design、slides）。" }
    },
    timeoutMs: 15_000,
    output: { schema: { type: "string" }, render: (_args, value) => [{ type: "text", text: value }] },
    execute: (args) => checkDeck(args)
  }));
}
