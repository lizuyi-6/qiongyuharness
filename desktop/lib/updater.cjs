"use strict";

const { spawn } = require("node:child_process");
const { createHash } = require("node:crypto");
// Electron 主进程会对 node:fs 的 .asar 路径做虚拟文件系统拦截，读写更新包时优先使用 original-fs 绕过拦截。
let rawFs;
try {
  rawFs = require("original-fs");
} catch {
  rawFs = require("node:fs");
}
const fs = rawFs;
const fsp = rawFs.promises;
const path = require("node:path");
const zlib = require("node:zlib");
const { sessionFetch } = require("./session-cookie.cjs");

const DELTA_FORMAT = "qo-delta-v1";
const MAX_PATCH_GZ_BYTES = 32 * 1024 * 1024;
const MAX_PATCH_JSON_BYTES = 64 * 1024 * 1024;
const ALLOWED_HARNESS_FILE_RE = /^qo-runtime\/harness\/[A-Za-z0-9._-]+\.(?:mjs|cjs|yml|yaml|json)$/;
const SAFE_FILENAME_RE = /^[A-Za-z0-9._-]+\.(?:qopatch|exe)$/;

function sha256Buffer(buffer) {
  return createHash("sha256").update(buffer).digest("hex");
}

function sha256File(filePath) {
  return new Promise((resolve, reject) => {
    const hash = createHash("sha256");
    const stream = fs.createReadStream(filePath);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("error", reject);
    stream.on("end", () => resolve(hash.digest("hex")));
  });
}

function computeRuntimeHash({ electronVersion, harnessLockPath }) {
  const electron = String(electronVersion || "").trim();
  if (!electron || !harnessLockPath) return "";
  try {
    const lockBytes = fs.readFileSync(harnessLockPath);
    const lockHash = sha256Buffer(lockBytes);
    return sha256Buffer(Buffer.from(`v1|${electron}|${lockHash}`, "utf8")).slice(0, 16);
  } catch {
    return "";
  }
}

function isAllowedPatchPath(relPath) {
  if (typeof relPath !== "string") return false;
  if (relPath === "app.asar") return true;
  return ALLOWED_HARNESS_FILE_RE.test(relPath);
}

function parseSemver(value) {
  const match = /^v?(\d+)\.(\d+)\.(\d+)/.exec(String(value || "").trim());
  if (!match) return null;
  return [Number(match[1]), Number(match[2]), Number(match[3])];
}

function isNewerVersion(latest, current) {
  const a = parseSemver(latest);
  const b = parseSemver(current);
  if (!a) return false;
  if (!b) return true;
  for (let i = 0; i < 3; i += 1) {
    if (a[i] !== b[i]) return a[i] > b[i];
  }
  return false;
}

function psQuote(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function buildDeltaApplyScript({ pid, resourcesDir, stagedDir, backupDir, exePath, appliedPath, errorLogPath, patchFile, files, version }) {
  const relList = files.map((file) => psQuote(file)).join(", ");
  return [
    "$ErrorActionPreference = 'Stop'",
    `$TargetPid = ${Number(pid)}`,
    `$ResourcesDir = ${psQuote(resourcesDir)}`,
    `$StagedDir = ${psQuote(stagedDir)}`,
    `$BackupDir = ${psQuote(backupDir)}`,
    `$ExePath = ${psQuote(exePath)}`,
    `$AppliedPath = ${psQuote(appliedPath)}`,
    `$ErrorLogPath = ${psQuote(errorLogPath)}`,
    `$PatchFile = ${psQuote(patchFile)}`,
    `$Version = ${psQuote(version)}`,
    `$Files = @(${relList})`,
    "",
    "# 等待旧版主进程完全退出以释放 app.asar 句柄",
    "for ($i = 0; $i -lt 100; $i++) {",
    "  if (-not (Get-Process -Id $TargetPid -ErrorAction SilentlyContinue)) { break }",
    "  Start-Sleep -Milliseconds 250",
    "}",
    "Start-Sleep -Milliseconds 350",
    "",
    "New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null",
    "$BackedUp = @()",
    "try {",
    "  foreach ($rel in $Files) {",
    "    $winRel = $rel -replace '/', '\\\\'",
    "    $src = Join-Path $StagedDir ($winRel + '.payload')",
    "    if (-not (Test-Path -LiteralPath $src)) { $src = Join-Path $StagedDir $winRel }",
    "    $dst = Join-Path $ResourcesDir $winRel",
    "    $bak = Join-Path $BackupDir $winRel",
    "    if (-not (Test-Path -LiteralPath $src)) { throw \"Missing staged file: $rel\" }",
    "    $dstParent = Split-Path -Parent $dst",
    "    if ($dstParent -and -not (Test-Path -LiteralPath $dstParent)) {",
    "      New-Item -ItemType Directory -Force -Path $dstParent | Out-Null",
    "    }",
    "    if (Test-Path -LiteralPath $dst) {",
    "      $bakParent = Split-Path -Parent $bak",
    "      if ($bakParent -and -not (Test-Path -LiteralPath $bakParent)) {",
    "        New-Item -ItemType Directory -Force -Path $bakParent | Out-Null",
    "      }",
    "      Copy-Item -LiteralPath $dst -Destination $bak -Force",
    "      $BackedUp += $rel",
    "    }",
    "    $copied = $false",
    "    for ($attempt = 0; $attempt -lt 20; $attempt++) {",
    "      try {",
    "        Copy-Item -LiteralPath $src -Destination $dst -Force",
    "        $copied = $true",
    "        break",
    "      } catch {",
    "        Start-Sleep -Milliseconds 300",
    "      }",
    "    }",
    "    if (-not $copied) { throw \"Failed to replace file: $rel\" }",
    "  }",
    "  $meta = @{ version = $Version; mode = 'delta'; applied_at = (Get-Date).ToUniversalTime().ToString('o') } | ConvertTo-Json -Compress",
    "  Set-Content -LiteralPath $AppliedPath -Value $meta -Encoding UTF8",
    "  Remove-Item -LiteralPath $StagedDir -Recurse -Force -ErrorAction SilentlyContinue",
    "  Remove-Item -LiteralPath $BackupDir -Recurse -Force -ErrorAction SilentlyContinue",
    "  if ($PatchFile -and (Test-Path -LiteralPath $PatchFile)) { Remove-Item -LiteralPath $PatchFile -Force -ErrorAction SilentlyContinue }",
    "  Remove-Item -LiteralPath $ErrorLogPath -Force -ErrorAction SilentlyContinue",
    "} catch {",
    "  $errMsg = $_.Exception.Message",
    "  Set-Content -LiteralPath $ErrorLogPath -Value \"$(Get-Date -Format o) $errMsg\" -Encoding UTF8",
    "  foreach ($rel in $BackedUp) {",
    "    try {",
    "      $winRel = $rel -replace '/', '\\\\'",
    "      $bak = Join-Path $BackupDir $winRel",
    "      $dst = Join-Path $ResourcesDir $winRel",
    "      if (Test-Path -LiteralPath $bak) { Copy-Item -LiteralPath $bak -Destination $dst -Force }",
    "    } catch {}",
    "  }",
    "}",
    "if (Test-Path -LiteralPath $ExePath) {",
    "  Start-Process -FilePath $ExePath",
    "}",
    ""
  ].join("\r\n");
}

function buildFullApplyScript({ pid, installerPath, exePath, appliedPath, version }) {
  return [
    "$ErrorActionPreference = 'Continue'",
    `$TargetPid = ${Number(pid)}`,
    `$InstallerPath = ${psQuote(installerPath)}`,
    `$ExePath = ${psQuote(exePath)}`,
    `$AppliedPath = ${psQuote(appliedPath)}`,
    `$Version = ${psQuote(version)}`,
    "",
    "for ($i = 0; $i -lt 100; $i++) {",
    "  if (-not (Get-Process -Id $TargetPid -ErrorAction SilentlyContinue)) { break }",
    "  Start-Sleep -Milliseconds 250",
    "}",
    "Start-Sleep -Milliseconds 350",
    "",
    "$proc = Start-Process -FilePath $InstallerPath -ArgumentList '/S' -Wait -PassThru",
    "if ($proc -and $proc.ExitCode -eq 0 -and (Test-Path -LiteralPath $ExePath)) {",
    "  $meta = @{ version = $Version; mode = 'full'; applied_at = (Get-Date).ToUniversalTime().ToString('o') } | ConvertTo-Json -Compress",
    "  Set-Content -LiteralPath $AppliedPath -Value $meta -Encoding UTF8",
    "  Remove-Item -LiteralPath $InstallerPath -Force -ErrorAction SilentlyContinue",
    "  Start-Process -FilePath $ExePath",
    "} elseif (Test-Path -LiteralPath $InstallerPath) {",
    "  Start-Process -FilePath $InstallerPath",
    "}",
    ""
  ].join("\r\n");
}

async function unpackAndVerifyDeltaPatch({ patchPath, stagedDir, expectedVersion, expectedRuntimeHash }) {
  const stat = await fsp.stat(patchPath);
  if (!stat.isFile() || stat.size <= 0 || stat.size > MAX_PATCH_GZ_BYTES) {
    throw new Error("增量更新包大小异常");
  }
  const gzBytes = await fsp.readFile(patchPath);
  let jsonBytes;
  try {
    jsonBytes = zlib.gunzipSync(gzBytes, { maxOutputLength: MAX_PATCH_JSON_BYTES });
  } catch {
    throw new Error("增量更新包解压失败，请重新下载");
  }
  let payload;
  try {
    payload = JSON.parse(jsonBytes.toString("utf8"));
  } catch {
    throw new Error("增量更新包元数据损坏");
  }
  if (!payload || typeof payload !== "object" || payload.format !== DELTA_FORMAT) {
    throw new Error("增量更新包格式不受支持");
  }
  if (expectedVersion && String(payload.version || "").trim() !== String(expectedVersion).trim()) {
    throw new Error("增量更新包版本与目标版本不符");
  }
  if (expectedRuntimeHash && String(payload.runtime_hash || "").trim().toLowerCase() !== String(expectedRuntimeHash).trim().toLowerCase()) {
    throw new Error("当前运行环境与增量包不匹配，请使用全量安装包更新");
  }
  if (!Array.isArray(payload.files) || payload.files.length === 0) {
    throw new Error("增量更新包缺少有效文件列表");
  }

  await fsp.rm(stagedDir, { recursive: true, force: true });
  await fsp.mkdir(stagedDir, { recursive: true });
  const stagedRoot = path.resolve(stagedDir);
  const relPaths = [];

  for (const item of payload.files) {
    if (!item || typeof item !== "object" || !isAllowedPatchPath(item.path)) {
      throw new Error(`增量更新包包含非法路径：${String(item?.path || "")}`);
    }
    if (typeof item.data !== "string" || typeof item.sha256 !== "string" || !Number.isSafeInteger(item.size) || item.size < 0) {
      throw new Error(`增量更新包文件条目无效：${item.path}`);
    }
    const content = Buffer.from(item.data, "base64");
    if (content.length !== item.size || sha256Buffer(content) !== item.sha256.toLowerCase()) {
      throw new Error(`增量更新包文件校验失败：${item.path}`);
    }
    const target = path.resolve(stagedRoot, `${item.path}.payload`);
    const relative = path.relative(stagedRoot, target);
    if (!relative || relative.startsWith("..") || path.isAbsolute(relative)) {
      throw new Error(`增量更新包路径越界：${item.path}`);
    }
    await fsp.mkdir(path.dirname(target), { recursive: true });
    await fsp.writeFile(target, content);
    relPaths.push(item.path);
  }

  if (!relPaths.includes("app.asar")) {
    throw new Error("增量更新包缺少核心应用文件 app.asar");
  }
  return { version: payload.version, files: relPaths };
}

function createDesktopUpdater({
  app,
  getWindow,
  getOrigin,
  quitApplication,
  emit = () => {},
  dependencies = {}
}) {
  const fetchWithSession = dependencies.sessionFetch || sessionFetch;
  const spawnImpl = dependencies.spawn || spawn;
  const isPackaged = dependencies.isPackaged !== undefined ? dependencies.isPackaged : Boolean(app?.isPackaged);
  const currentVersion = dependencies.currentVersion || (typeof app?.getVersion === "function" ? app.getVersion() : "0.0.0");
  const updatesDir = dependencies.updatesDir || path.join(app.getPath("userData"), "updates");
  const resourcesDir = dependencies.resourcesDir || ((isPackaged && process.resourcesPath) ? process.resourcesPath : path.resolve(__dirname, ".."));
  const exePath = dependencies.exePath || process.execPath;
  const harnessLockPath = dependencies.harnessLockPath || (
    isPackaged
      ? path.join(resourcesDir, "qo-runtime", "harness", "package-lock.json")
      : path.resolve(__dirname, "..", "..", "harness", "package-lock.json")
  );
  const runtimeHash = dependencies.runtimeHash !== undefined
    ? dependencies.runtimeHash
    : computeRuntimeHash({
      electronVersion: dependencies.electronVersion || process.versions.electron,
      harnessLockPath
    });

  let activeAbort = null;
  let readyArtifact = null;
  let lastManifest = null;
  let appliedNotice = null;
  let state = {
    status: "idle",
    currentVersion,
    latestVersion: "",
    mode: "none",
    notes: "",
    publishedAt: "",
    size: 0,
    fullSize: 0,
    deltaSize: 0,
    hasDelta: false,
    hasFull: false,
    received: 0,
    total: 0,
    percent: 0,
    bytesPerSecond: 0,
    error: "",
    checkedAt: "",
    canInstallInApp: process.platform === "win32" && isPackaged,
    runtimeHash,
    appliedNotice: null
  };

  function readAppliedNoticeOnce() {
    const appliedPath = path.join(updatesDir, "last-applied.json");
    try {
      if (!fs.existsSync(appliedPath)) return null;
      const raw = JSON.parse(fs.readFileSync(appliedPath, "utf8"));
      fs.rmSync(appliedPath, { force: true });
      if (raw && typeof raw.version === "string" && raw.version.trim() === currentVersion) {
        return { version: raw.version.trim(), mode: raw.mode === "full" ? "full" : "delta", appliedAt: raw.applied_at || "" };
      }
    } catch {}
    return null;
  }

  appliedNotice = readAppliedNoticeOnce();
  if (appliedNotice) state.appliedNotice = appliedNotice;

  function snapshot() {
    return { ...state };
  }

  function updateState(patch) {
    state = { ...state, ...patch };
    emit(snapshot());
    return snapshot();
  }

  function currentSession() {
    if (dependencies.session) return dependencies.session;
    const win = getWindow?.();
    if (!win || win.isDestroyed()) throw new Error("桌面窗口已关闭");
    return win.webContents.session;
  }

  async function checkForUpdates({ silent = false } = {}) {
    if (["downloading", "verifying", "installing"].includes(state.status)) {
      return snapshot();
    }
    const origin = getOrigin?.();
    if (!origin) {
      if (!silent) updateState({ status: "error", error: "尚未连接工作台服务" });
      return snapshot();
    }
    if (!silent) {
      updateState({ status: "checking", error: "" });
    }
    try {
      const url = new URL("/api/v1/updates/latest", origin);
      url.searchParams.set("current", currentVersion);
      if (runtimeHash) url.searchParams.set("runtime_hash", runtimeHash);
      const response = await fetchWithSession(currentSession(), origin, url.toString(), {
        method: "GET",
        headers: { Accept: "application/json" },
        redirect: "error",
        signal: AbortSignal.timeout(15_000)
      });
      if (response.status === 401) {
        if (!silent) updateState({ status: "error", error: "请先登录后检查更新" });
        return snapshot();
      }
      if (!response.ok) {
        throw new Error(`检查更新失败（HTTP ${response.status}）`);
      }
      const data = await response.json();
      lastManifest = data;
      const latestVersion = String(data?.latest_version || "").trim();
      const delta = data?.delta || null;
      const full = data?.full || null;
      const hasUpdate = Boolean(data?.has_update) && isNewerVersion(latestVersion, currentVersion) && Boolean(delta || full);
      const mode = hasUpdate ? (data?.update_mode === "delta" && delta ? "delta" : full ? "full" : delta ? "delta" : "none") : "none";
      const targetPkg = mode === "delta" ? delta : full;

      // 如果已经下载并校验好同一个版本且模式相同，保持 ready 状态
      if (hasUpdate && readyArtifact && readyArtifact.version === latestVersion && readyArtifact.mode === mode) {
        return updateState({
          status: "ready",
          latestVersion,
          mode,
          notes: String(data?.notes || ""),
          publishedAt: String(data?.published_at || ""),
          size: targetPkg?.size || 0,
          fullSize: full?.size || 0,
          deltaSize: delta?.size || 0,
          hasDelta: Boolean(delta),
          hasFull: Boolean(full),
          error: "",
          checkedAt: new Date().toISOString()
        });
      }

      if (!hasUpdate) {
        return updateState({
          status: "up-to-date",
          latestVersion: latestVersion || currentVersion,
          mode: "none",
          notes: String(data?.notes || ""),
          publishedAt: String(data?.published_at || ""),
          size: 0,
          fullSize: full?.size || 0,
          deltaSize: delta?.size || 0,
          hasDelta: Boolean(delta),
          hasFull: Boolean(full),
          received: 0,
          total: 0,
          percent: 0,
          bytesPerSecond: 0,
          error: "",
          checkedAt: new Date().toISOString()
        });
      }

      return updateState({
        status: "available",
        latestVersion,
        mode,
        notes: String(data?.notes || ""),
        publishedAt: String(data?.published_at || ""),
        size: targetPkg?.size || 0,
        fullSize: full?.size || 0,
        deltaSize: delta?.size || 0,
        hasDelta: Boolean(delta),
        hasFull: Boolean(full),
        received: 0,
        total: targetPkg?.size || 0,
        percent: 0,
        bytesPerSecond: 0,
        error: "",
        checkedAt: new Date().toISOString()
      });
    } catch (error) {
      if (silent) return snapshot();
      return updateState({
        status: "error",
        error: error?.message || "检查更新失败，请稍后重试"
      });
    }
  }

  function cancelDownload() {
    if (activeAbort && !activeAbort.signal.aborted) {
      activeAbort.abort();
      return true;
    }
    return false;
  }

  async function downloadUpdate({ mode: requestedMode } = {}) {
    if (["downloading", "verifying", "installing"].includes(state.status)) {
      return snapshot();
    }
    const origin = getOrigin?.();
    if (!origin) throw new Error("尚未连接工作台服务");
    if (!lastManifest || !lastManifest.has_update) {
      await checkForUpdates({ silent: false });
    }
    if (!lastManifest || !lastManifest.has_update) {
      return snapshot();
    }

    const latestVersion = String(lastManifest.latest_version || "").trim();
    const delta = lastManifest.delta || null;
    const full = lastManifest.full || null;
    let mode = requestedMode === "full" ? "full" : requestedMode === "delta" ? "delta" : lastManifest.update_mode;
    if (mode === "delta" && !delta) mode = "full";
    if (mode === "full" && !full) mode = "delta";
    const pkg = mode === "delta" ? delta : full;
    if (!pkg || !pkg.file || !SAFE_FILENAME_RE.test(pkg.file) || !pkg.url) {
      return updateState({ status: "error", error: "服务端未提供可下载的更新包" });
    }

    const downloadUrl = new URL(pkg.url, origin);
    if (downloadUrl.origin !== new URL(origin).origin || !downloadUrl.pathname.startsWith("/api/v1/updates/files/")) {
      return updateState({ status: "error", error: "更新包下载地址不合法" });
    }

    await fsp.mkdir(updatesDir, { recursive: true });
    const finalPath = path.join(updatesDir, pkg.file);
    const partPath = `${finalPath}.part`;
    const expectedSha = String(pkg.sha256 || "").trim().toLowerCase();
    const expectedSize = Number(pkg.size) || 0;

    // 命中本地已下载完成的包时跳过下载直接验签
    let reuseFinal = false;
    if (expectedSha && fs.existsSync(finalPath)) {
      try {
        const existingSha = await sha256File(finalPath);
        if (existingSha === expectedSha) reuseFinal = true;
      } catch {}
    }

    const controller = new AbortController();
    activeAbort = controller;
    try {
      if (!reuseFinal) {
        let existingBytes = 0;
        try {
          const st = await fsp.stat(partPath);
          if (st.isFile() && st.size > 0 && (!expectedSize || st.size < expectedSize)) {
            existingBytes = st.size;
          } else {
            await fsp.rm(partPath, { force: true });
          }
        } catch {}

        updateState({
          status: "downloading",
          mode,
          size: expectedSize,
          received: existingBytes,
          total: expectedSize,
          percent: expectedSize > 0 ? Math.min(99, Math.round((existingBytes / expectedSize) * 100)) : 0,
          bytesPerSecond: 0,
          error: ""
        });

        const headers = {};
        if (existingBytes > 0) headers.Range = `bytes=${existingBytes}-`;
        const response = await fetchWithSession(currentSession(), origin, downloadUrl.toString(), {
          method: "GET",
          headers,
          redirect: "error",
          signal: controller.signal
        });
        if (!response.ok && response.status !== 206) {
          throw new Error(`下载更新包失败（HTTP ${response.status}）`);
        }

        const isResumed = response.status === 206 && existingBytes > 0;
        const startOffset = isResumed ? existingBytes : 0;
        const contentLength = Number(response.headers.get("content-length")) || 0;
        const total = expectedSize || (contentLength ? startOffset + contentLength : 0);
        let received = startOffset;

        const fileStream = fs.createWriteStream(partPath, { flags: isResumed ? "a" : "w" });
        const reader = response.body?.getReader?.();
        if (!reader) {
          fileStream.close();
          throw new Error("下载响应流不可用");
        }

        let lastTickAt = Date.now();
        let lastTickBytes = received;
        let speed = 0;
        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            if (!value || value.byteLength === 0) continue;
            await new Promise((resolve, reject) => {
              if (!fileStream.write(Buffer.from(value))) {
                fileStream.once("drain", resolve);
                fileStream.once("error", reject);
              } else {
                resolve();
              }
            });
            received += value.byteLength;
            const now = Date.now();
            const elapsed = now - lastTickAt;
            if (elapsed >= 220 || (total > 0 && received >= total)) {
              if (elapsed > 0) {
                speed = Math.round(((received - lastTickBytes) * 1000) / elapsed);
              }
              lastTickAt = now;
              lastTickBytes = received;
              updateState({
                status: "downloading",
                mode,
                received,
                total: total || received,
                percent: total > 0 ? Math.min(99, Math.round((received / total) * 100)) : 0,
                bytesPerSecond: speed
              });
            }
          }
          await new Promise((resolve, reject) => {
            fileStream.on("error", reject);
            fileStream.end(resolve);
          });
        } catch (err) {
          fileStream.destroy();
          throw err;
        }

        await fsp.rm(finalPath, { force: true });
        await fsp.rename(partPath, finalPath);
      }

      updateState({ status: "verifying", mode, percent: 100, bytesPerSecond: 0, error: "" });
      if (expectedSha) {
        const actualSha = await sha256File(finalPath);
        if (actualSha !== expectedSha) {
          await fsp.rm(finalPath, { force: true });
          throw new Error("更新包完整性校验失败（SHA-256 不一致），请重新下载");
        }
      }

      if (mode === "delta") {
        const stagedDir = path.join(updatesDir, `staged-${latestVersion}`);
        const unpacked = await unpackAndVerifyDeltaPatch({
          patchPath: finalPath,
          stagedDir,
          expectedVersion: latestVersion,
          expectedRuntimeHash: runtimeHash
        });
        readyArtifact = {
          mode: "delta",
          version: latestVersion,
          patchPath: finalPath,
          stagedDir,
          files: unpacked.files
        };
      } else {
        readyArtifact = {
          mode: "full",
          version: latestVersion,
          installerPath: finalPath
        };
      }

      return updateState({
        status: "ready",
        latestVersion,
        mode,
        size: expectedSize || state.size,
        received: expectedSize || state.received,
        total: expectedSize || state.total,
        percent: 100,
        bytesPerSecond: 0,
        error: ""
      });
    } catch (error) {
      if (controller.signal.aborted || error?.name === "AbortError") {
        return updateState({
          status: "available",
          bytesPerSecond: 0,
          error: ""
        });
      }
      return updateState({
        status: "error",
        bytesPerSecond: 0,
        error: error?.message || "下载更新失败，请重试"
      });
    } finally {
      if (activeAbort === controller) activeAbort = null;
    }
  }

  async function installUpdate() {
    if (state.status !== "ready" || !readyArtifact) {
      throw new Error("尚未准备好可安装的更新包");
    }
    if (process.platform !== "win32") {
      throw new Error("当前平台暂不支持自动应用更新，请手动下载安装");
    }
    if (!isPackaged) {
      throw new Error("开发运行模式不支持覆盖安装，请使用已安装的桌面客户端测试更新");
    }

    updateState({ status: "installing", error: "" });
    const appliedPath = path.join(updatesDir, "last-applied.json");
    const errorLogPath = path.join(updatesDir, "last-error.log");
    const scriptPath = path.join(updatesDir, `apply-${readyArtifact.mode}-${readyArtifact.version}.ps1`);

    let scriptContent;
    if (readyArtifact.mode === "delta") {
      const backupDir = path.join(updatesDir, `backup-${readyArtifact.version}`);
      scriptContent = buildDeltaApplyScript({
        pid: process.pid,
        resourcesDir,
        stagedDir: readyArtifact.stagedDir,
        backupDir,
        exePath,
        appliedPath,
        errorLogPath,
        patchFile: readyArtifact.patchPath,
        files: readyArtifact.files,
        version: readyArtifact.version
      });
    } else {
      scriptContent = buildFullApplyScript({
        pid: process.pid,
        installerPath: readyArtifact.installerPath,
        exePath,
        appliedPath,
        version: readyArtifact.version
      });
    }

    // 写入 UTF-8 BOM，确保 Windows PowerShell 5.1 正确解析中文路径（如 穹驭办公.exe）
    await fsp.writeFile(scriptPath, `\uFEFF${scriptContent}`, { encoding: "utf8" });

    const powershellExe = path.join(process.env.SystemRoot || "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
    const child = spawnImpl(powershellExe, [
      "-NoProfile",
      "-NonInteractive",
      "-ExecutionPolicy", "Bypass",
      "-WindowStyle", "Hidden",
      "-File", scriptPath
    ], {
      detached: true,
      windowsHide: true,
      stdio: "ignore"
    });
    child.unref?.();

    setTimeout(() => {
      quitApplication?.();
    }, 180);
    return snapshot();
  }

  function clearAppliedNotice() {
    if (!state.appliedNotice) return snapshot();
    return updateState({ appliedNotice: null });
  }

  return {
    getState: snapshot,
    checkForUpdates,
    downloadUpdate,
    cancelDownload,
    installUpdate,
    clearAppliedNotice
  };
}

module.exports = {
  DELTA_FORMAT,
  buildDeltaApplyScript,
  buildFullApplyScript,
  computeRuntimeHash,
  createDesktopUpdater,
  isAllowedPatchPath,
  isNewerVersion,
  sha256Buffer,
  sha256File,
  unpackAndVerifyDeltaPatch
};
