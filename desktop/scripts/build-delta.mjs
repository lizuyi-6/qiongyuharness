import { createHash } from "node:crypto";
import { mkdir, readdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import zlib from "node:zlib";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const desktopDir = path.resolve(scriptDir, "..");
const repoDir = path.resolve(desktopDir, "..");
const distDir = path.join(desktopDir, "dist");
const sourceHarnessDir = path.join(repoDir, "harness");

function sha256(buffer) {
  return createHash("sha256").update(buffer).digest("hex");
}

async function resolveElectronVersion(pkg) {
  try {
    const installed = JSON.parse(await readFile(path.join(desktopDir, "node_modules", "electron", "package.json"), "utf8"));
    if (typeof installed.version === "string" && installed.version.trim()) {
      return installed.version.trim();
    }
  } catch {}
  return String(pkg?.devDependencies?.electron || "").replace(/^[~^]/, "").trim();
}

const pkg = JSON.parse(await readFile(path.join(desktopDir, "package.json"), "utf8"));
const version = String(pkg.version || "").trim();
if (!version) throw new Error("desktop/package.json 缺少 version");

const electronVersion = await resolveElectronVersion(pkg);
if (!electronVersion) throw new Error("无法确定 Electron 版本");

const lockBytes = await readFile(path.join(sourceHarnessDir, "package-lock.json"));
const lockHash = sha256(lockBytes);
const runtimeHash = sha256(Buffer.from(`v1|${electronVersion}|${lockHash}`, "utf8")).slice(0, 16);

const asarPath = path.join(distDir, "win-unpacked", "resources", "app.asar");
let asarBytes;
try {
  asarBytes = await readFile(asarPath);
} catch {
  throw new Error(`未找到 ${asarPath}，请先运行 electron-builder --win --x64（或 --dir）`);
}

const files = [
  {
    path: "app.asar",
    size: asarBytes.length,
    sha256: sha256(asarBytes),
    data: asarBytes.toString("base64")
  }
];

const harnessEntries = await readdir(sourceHarnessDir);
const harnessFiles = harnessEntries
  .filter((name) => name === "package.json" || name === "office.patch.yml" || name.endsWith(".mjs"))
  .sort((a, b) => a.localeCompare(b));

for (const name of harnessFiles) {
  const bytes = await readFile(path.join(sourceHarnessDir, name));
  files.push({
    path: `qo-runtime/harness/${name}`,
    size: bytes.length,
    sha256: sha256(bytes),
    data: bytes.toString("base64")
  });
}

const patchPayload = {
  format: "qo-delta-v1",
  version,
  runtime_hash: runtimeHash,
  electron_version: electronVersion,
  created_at: new Date().toISOString(),
  files
};

const rawJson = Buffer.from(JSON.stringify(patchPayload), "utf8");
const gzBytes = zlib.gzipSync(rawJson, { level: 9 });
const deltaFile = `QiongyuOffice-${version}-delta.qopatch`;
const deltaPath = path.join(distDir, deltaFile);
await mkdir(distDir, { recursive: true });
await writeFile(deltaPath, gzBytes);
const deltaSha = sha256(gzBytes);

let fullInfo = null;
const fullFile = `QiongyuOffice-${version}-x64.exe`;
const fullPath = path.join(distDir, fullFile);
try {
  const st = await stat(fullPath);
  if (st.isFile() && st.size > 0) {
    const fullBytes = await readFile(fullPath);
    fullInfo = {
      file: fullFile,
      size: st.size,
      sha256: sha256(fullBytes)
    };
  }
} catch {}

if (!fullInfo) {
  try {
    const prevManifest = JSON.parse(await readFile(path.join(distDir, "manifest.json"), "utf8"));
    if (prevManifest?.full?.file && prevManifest?.full?.sha256 && Number(prevManifest?.full?.size) > 0) {
      fullInfo = prevManifest.full;
    }
  } catch {}
}

const notes = String(process.env.QO_RELEASE_NOTES || "").trim() ||
  `穹驭办公 v${version}：支持应用内秒级增量更新与在线下载安装。`;

const manifest = {
  version,
  published_at: new Date().toISOString(),
  notes,
  runtime_hash: runtimeHash,
  min_delta_version: "0.2.12",
  delta: {
    file: deltaFile,
    size: gzBytes.length,
    sha256: deltaSha
  },
  full: fullInfo
};

const manifestPath = path.join(distDir, "manifest.json");
await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");

console.log(JSON.stringify({
  ok: true,
  version,
  electronVersion,
  runtimeHash,
  filesIncluded: files.map((f) => f.path),
  delta: manifest.delta,
  full: manifest.full
}, null, 2));
