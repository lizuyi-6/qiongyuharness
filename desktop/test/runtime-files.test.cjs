const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const REPO_ROOT = path.resolve(__dirname, "..", "..");

function stagedHarnessFiles() {
  const script = fs.readFileSync(path.join(REPO_ROOT, "desktop", "scripts", "prepare-runtime.mjs"), "utf8");
  const match = script.match(/const HARNESS_FILES = (\[[^\]]*\]);/);
  assert.ok(match, "prepare-runtime.mjs 必须用 HARNESS_FILES 声明运行时清单");
  return JSON.parse(match[1]);
}

test("安装包运行时清单包含 harness/ 下全部插件与桌面端引用的文件", () => {
  const staged = stagedHarnessFiles();
  const plugins = fs.readdirSync(path.join(REPO_ROOT, "harness")).filter((name) => name.endsWith(".mjs"));
  assert.ok(plugins.includes("office-image.mjs"));
  for (const plugin of plugins) assert.ok(staged.includes(plugin), `${plugin} 未进入安装包运行时清单`);

  const desktopSources = ["agent.cjs", path.join("lib", "runtime-env.cjs")]
    .map((file) => fs.readFileSync(path.join(REPO_ROOT, "desktop", file), "utf8")).join("\n");
  // 只看从安装目录 harnessRoot 直接加载的插件/配置文件；node_modules 由 npm ci 装入，
  // 每轮生成在临时目录里的 *.patch.yml 不随包分发。
  const referenced = [...new Set([...desktopSources.matchAll(/harnessRoot,\s*"([^"/]+\.(?:mjs|yml|json))"/g)]
    .map((match) => match[1]))];
  assert.ok(referenced.includes("office-image.mjs") && referenced.includes("office.patch.yml"));
  for (const file of referenced) assert.ok(staged.includes(file), `${file} 被桌面端加载但未进入安装包运行时清单`);
});
