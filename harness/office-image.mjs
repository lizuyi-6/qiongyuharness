import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const resolve = createRequire(process.env.QO_DSH_EXECUTABLE || import.meta.url);
const { defineTool } = await import(pathToFileURL(resolve.resolve("@deepseek-ai/dsh-tools")).href);
const ALLOWED_TOOLS = new Set(["office_context_read", "office_image_generate", "office_ppt_check"]);
const SIZES = ["1024x1024", "1536x1024", "1024x1536", "1536x864"];
const QUALITIES = ["low", "medium", "high"];

export const name = "qiongyu-office-image";
export const inject = ["tools"];

function brokerAddress() {
  const raw = process.env.QO_IMAGE_BROKER_URL;
  const key = process.env.QO_IMAGE_BROKER_KEY;
  if (!raw || !key) throw new Error("图片服务不可用");
  const url = new URL(raw);
  if (url.protocol !== "http:" || url.hostname !== "127.0.0.1" || url.username || url.password || !["", "/"].includes(url.pathname)) {
    throw new Error("图片服务地址无效");
  }
  return { url: `${raw.replace(/\/+$/, "")}/generate`, key };
}

async function generate(args, execution) {
  const prompt = typeof args.prompt === "string" ? args.prompt.trim() : "";
  const size = args.size || "1024x1024";
  const quality = args.quality || "medium";
  if (prompt.length < 1 || prompt.length > 4000 || !SIZES.includes(size) || !QUALITIES.includes(quality)) {
    return "生图失败：生图请求参数无效";
  }
  let endpoint;
  try {
    const broker = brokerAddress();
    endpoint = broker.url;
    const response = await fetch(endpoint, {
      method: "POST",
      headers: { Authorization: `Bearer ${broker.key}`, "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify({ prompt, size, quality }),
      signal: execution.signal,
      redirect: "error"
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || result?.ok !== true) {
      const message = typeof result?.message === "string" ? result.message.slice(0, 300) : `图片服务返回 HTTP ${response.status}`;
      return `生图失败：${message}`;
    }
    if (!Number.isSafeInteger(result.index) || result.index < 1 || !SIZES.includes(result.size)) {
      return "生图失败：图片服务返回内容无效";
    }
    return `已生成第 ${result.index} 张图片（image_index=${result.index}，${result.size}，${qualityLabel(quality)}质量）。若是制作 PPT，可直接把 ${result.index} 填入 deck.slides 的 bg_image 或 {"type":"image","image_index":${result.index},...}；不要在 reply 中输出图片链接、Markdown 图片或 base64。`;
  } catch (error) {
    const message = error?.name === "AbortError" ? "工具调用已取消" : "图片服务暂不可用";
    return `生图失败：${message}`;
  }
}

function qualityLabel(quality) {
  return quality === "low" ? "草稿" : quality === "high" ? "精细" : "标准";
}

export function apply(ctx) {
  ctx.tools.guard((execution) => ALLOWED_TOOLS.has(execution.name) ? undefined : "本轮只允许调用办公资料读取和生图工具");
  ctx.tools.register(defineTool({
    name: "office_image_generate",
    description: "用于生成 PPT 封面主视觉/配图，或用户明确需要的插图、海报、示意图等；按张计费，PPT 配图提示词中请要求画面无文字。",
    parameters: {
      prompt: { type: "string", required: true, description: "图片生成提示词，1 至 4000 字。" },
      size: { type: "string", enum: SIZES, default: "1024x1024", description: "图片尺寸。" },
      quality: { type: "string", enum: QUALITIES, default: "medium", description: "图片质量。" }
    },
    timeoutMs: 120_000,
    output: { schema: { type: "string" }, render: (_args, value) => [{ type: "text", text: value }] },
    execute: generate
  }));
}
