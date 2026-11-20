import { spawnSync } from "node:child_process";
import { cp, mkdir, readdir, readFile, rm } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const desktopDirectory = path.resolve(scriptDirectory, "..");
const repositoryDirectory = path.resolve(desktopDirectory, "..");
const sourceHarness = path.join(repositoryDirectory, "harness");
const runtimeStage = path.join(desktopDirectory, "runtime-stage");
const runtimeRoot = path.join(runtimeStage, "qo-runtime");
const harnessRoot = path.join(runtimeRoot, "harness");

function atLeast(version, minimum = [24, 9, 0]) {
  const numbers = String(version).replace(/^v/, "").split(".").map((part) => Number.parseInt(part, 10));
  for (let index = 0; index < minimum.length; index += 1) {
    if ((numbers[index] || 0) > minimum[index]) return true;
    if ((numbers[index] || 0) < minimum[index]) return false;
  }
  return true;
}

function assertInside(target, root) {
  const resolvedTarget = path.resolve(target);
  const resolvedRoot = path.resolve(root);
  const relative = path.relative(resolvedRoot, resolvedTarget);
  if (!relative || relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new Error("运行时暂存目录不在桌面任务工作区内");
  }
}

if (process.arch !== "x64") throw new Error(`运行时构建要求 x64 Node，当前为 ${process.arch}`);
if (!atLeast(process.versions.node)) throw new Error(`Node ${process.versions.node} 低于要求的 24.9.0`);
assertInside(runtimeRoot, runtimeStage);
await rm(runtimeRoot, { recursive: true, force: true });

// 运行时只带这些文件；agent.cjs 会按需生成 patch 加载其中的 .mjs 插件。
const HARNESS_FILES = ["package.json", "package-lock.json", "office.patch.yml", "office-context.mjs", "office-image.mjs"];
// 漏带插件会让对应功能在安装包里整轮失败，所以 harness/ 下任何 .mjs 不在清单里就直接中止打包。
const missingPlugins = (await readdir(sourceHarness)).filter((name) => name.endsWith(".mjs") && !HARNESS_FILES.includes(name));
if (missingPlugins.length) throw new Error(`harness 插件未列入运行时清单：${missingPlugins.join(", ")}`);

for (const file of HARNESS_FILES) {
  const source = path.join(sourceHarness, file);
  await readFile(source);
  await mkdir(harnessRoot, { recursive: true });
  await cp(source, path.join(harnessRoot, file));
}

const nodeDirectory = path.join(runtimeRoot, "node");
await mkdir(nodeDirectory, { recursive: true });
await cp(process.execPath, path.join(nodeDirectory, "node.exe"));

assertInside(path.join(harnessRoot, "node_modules"), runtimeRoot);

const npmExecPath = process.env.npm_execpath;
const command = npmExecPath ? process.execPath : process.platform === "win32" ? (process.env.ComSpec || "cmd.exe") : "npm";
const args = npmExecPath
  ? [npmExecPath, "ci", "--omit=dev", "--ignore-scripts"]
  : process.platform === "win32"
    ? ["/d", "/s", "/c", "npm.cmd ci --omit=dev --ignore-scripts"]
    : ["ci", "--omit=dev", "--ignore-scripts"];
const install = spawnSync(command, args, { cwd: harnessRoot, stdio: "inherit", windowsHide: true });
if (install.error) throw install.error;
if (install.status !== 0) throw new Error(`Harness 运行时依赖安装失败，退出码 ${install.status}`);

const nodeCheck = spawnSync(path.join(nodeDirectory, "node.exe"), ["--version"], {
  encoding: "utf8", windowsHide: true
});
if (nodeCheck.status !== 0 || !atLeast(nodeCheck.stdout.trim())) {
  throw new Error("暂存的 node.exe 版本校验失败");
}
console.log(`Runtime staged: Node ${nodeCheck.stdout.trim()}, Harness ${harnessRoot}`);
