import { readFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const resolve = createRequire(process.env.QO_DSH_EXECUTABLE || import.meta.url);
const { defineTool } = await import(pathToFileURL(resolve.resolve("@deepseek-ai/dsh-tools")).href);

export const name = "qiongyu-office-context";
export const inject = ["tools"];

export function apply(ctx) {
  ctx.tools.register(defineTool({
    name: "office_context_read",
    description: "读取本轮已授权的项目资料和选中幻灯片的可编辑文本，参数不含文件路径。",
    parameters: {},
    output: {
      schema: { type: "string" },
      render: (_args, value) => [{ type: "text", text: value }],
    },
    async execute(_args, execution) {
      return readFile(process.env.QO_DSH_PLUGIN_CONTEXT, { encoding: "utf8", signal: execution.signal });
    },
  }));
}
