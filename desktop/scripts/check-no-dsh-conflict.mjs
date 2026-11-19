import { readFile, readdir, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const desktopDirectory = path.resolve(scriptDirectory, "..");
const resourcesDirectory = path.join(desktopDirectory, "dist", "win-unpacked", "resources");
const packageJson = JSON.parse(await readFile(path.join(desktopDirectory, "package.json"), "utf8"));
const installer = await readFile(path.join(desktopDirectory, "installer.nsi"), "utf8");
const runtimeEnv = await readFile(path.join(desktopDirectory, "lib", "runtime-env.cjs"), "utf8");
const forbiddenNames = /deepseek harness|dsh desktop|deepseek-harness|dsh-desktop|dsh-plugin-desktop|dsh-desktop-next/i;
const violations = [];
let checked = 0;

function check(condition, message) {
  checked += 1;
  if (!condition) violations.push(message);
}

function forbiddenArchivePath(value) {
  const normalized = value.replaceAll("\\", "/").toLowerCase();
  return normalized.includes("node_modules/@deepseek-ai/") || normalized.includes("dsh/node_modules/@deepseek-ai/");
}

async function archiveEntries(archive) {
  const handle = await import("node:fs/promises").then(({ open }) => open(archive, "r"));
  try {
    const prefix = Buffer.alloc(16);
    await handle.read(prefix, 0, prefix.length, 0);
    if (prefix.readUInt32LE(0) !== 4) throw new Error("invalid ASAR header");
    const headerSize = prefix.readUInt32LE(12);
    if (headerSize < 1 || headerSize > 32 * 1024 * 1024) throw new Error("ASAR header size invalid");
    const headerData = Buffer.alloc(headerSize);
    await handle.read(headerData, 0, headerData.length, 16);
    const header = JSON.parse(headerData.toString("utf8"));
    const entries = [];
    function walk(node, parent = "") {
      for (const [name, entry] of Object.entries(node.files || {})) {
        const relative = parent ? `${parent}/${name}` : name;
        entries.push(relative);
        if (entry.files) walk(entry, relative);
      }
    }
    walk(header);
    return entries;
  } finally {
    await handle.close();
  }
}

async function directoryEntries(root) {
  const entries = [];
  async function walk(current, relative = "") {
    for (const item of await readdir(current, { withFileTypes: true })) {
      const childRelative = relative ? `${relative}/${item.name}` : item.name;
      entries.push(childRelative);
      if (item.isDirectory()) await walk(path.join(current, item.name), childRelative);
    }
  }
  await walk(root);
  return entries;
}

check(typeof packageJson.name === "string" && !forbiddenNames.test(packageJson.name), "根 package.json name 含冲突名称");
check(typeof packageJson.build?.productName === "string" && !forbiddenNames.test(packageJson.build.productName), "安装显示名含冲突名称");
check(typeof packageJson.build?.artifactName === "string" && !forbiddenNames.test(packageJson.build.artifactName), "exe/安装包文件名含冲突名称");
check(!forbiddenNames.test(packageJson.build?.win?.executableName || packageJson.build?.productName || ""), "exe 名含冲突名称");
check(!forbiddenNames.test(installer), "installer.nsi 含冲突显示名、快捷方式名或安装目录");
check(installer.includes('CreateShortcut "$SMPROGRAMS\\穹驭办公.lnk"') && installer.includes('"DisplayName" "穹驭办公"'),
  "快捷方式名或卸载注册表 DisplayName 未使用穹驭办公");
check(installer.includes('RMDir /r "$LOCALAPPDATA\\Programs\\QiongyuOffice"') && installer.includes('RMDir /r "$APPDATA\\QiongyuOffice"'),
  "卸载范围未限定为穹驭办公安装目录和 userData");
check(Array.isArray(packageJson.build?.files) && packageJson.build.files.includes("agent.cjs") && packageJson.build.files.includes("lib/**"),
  "Electron files 未显式包含桌面主进程逻辑");
check(!packageJson.build.files.some((entry) => /harness|runtime-stage|node_modules/i.test(entry)),
  "Electron files 把 Harness 或暂存运行时纳入 app.asar");
check(packageJson.build?.extraResources?.some((entry) => entry.from === "runtime-stage/qo-runtime" && entry.to === "qo-runtime"),
  "Harness 运行时未配置到 resources/qo-runtime");
check(runtimeEnv.includes('env.HOME = home') && runtimeEnv.includes('env.DSH_HOME = home') && runtimeEnv.includes('env.DSH_AGENTS_HOME = home'),
  "DSH 用户目录未全部重定向到本轮临时目录");
check(runtimeEnv.includes('"sdk-minimal"') && !runtimeEnv.includes('"desktop"'), "DSH profile 不符合约定");

const configOnly = process.argv.includes("--config-only");
let distributionPresent = false;
try { distributionPresent = (await stat(resourcesDirectory)).isDirectory(); } catch {}
if (!configOnly && distributionPresent) {
  const runtimeDsh = path.join(resourcesDirectory, "qo-runtime", "harness", "node_modules", "@deepseek-ai", "dsh", "lib", "bin.js");
  try { await stat(runtimeDsh); check(true, ""); } catch { check(false, "resources/qo-runtime 中缺少已打包的 DSH CLI"); }
  for (const layout of ["app", "app.asar.unpacked"]) {
    const root = path.join(resourcesDirectory, layout);
    try {
      const entries = await directoryEntries(root);
      check(entries.every((entry) => !forbiddenArchivePath(entry)), `${layout} 含有 DSH 冲突路径`);
    } catch (error) {
      if (error.code !== "ENOENT") violations.push(`${layout} 检查失败: ${error.message}`);
    }
  }
  try {
    const entries = await archiveEntries(path.join(resourcesDirectory, "app.asar"));
    check(entries.every((entry) => !forbiddenArchivePath(entry)), "resources/app.asar 含有 DSH 冲突路径");
  } catch (error) {
    if (error.code !== "ENOENT") violations.push(`app.asar 检查失败: ${error.message}`);
  }
}

const mode = !configOnly && distributionPresent ? "dist/win-unpacked" : "config-only";
if (violations.length) {
  console.error(JSON.stringify({ mode, checked, violations }, null, 2));
  process.exitCode = 1;
} else {
  console.log(JSON.stringify({ mode, checked, violations: [], result: "pass" }));
}
