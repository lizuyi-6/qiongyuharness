"use strict";

const assert = require("node:assert/strict");
const { spawn } = require("node:child_process");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const {
  buildDeltaNodeScript,
  psEncodedCommand
} = require("../lib/updater.cjs");

test("updater.cjs does not use ASCII VBS launcher", async () => {
  const src = await fsp.readFile(path.join(__dirname, "..", "lib", "updater.cjs"), "utf8");
  assert.doesNotMatch(src, /encoding:\s*["']ascii["']/i);
  assert.doesNotMatch(src, /WScript\.Shell/i);
});

test("psEncodedCommand preserves Chinese username and exe path in UTF-16LE", () => {
  const sample = "Start-Process -FilePath 'C:\\Users\\张伟\\AppData\\Local\\Programs\\穹驭办公.exe'";
  const encoded = psEncodedCommand(sample);
  assert.match(encoded, /^[A-Za-z0-9+/=]+$/);
  const decoded = Buffer.from(encoded, "base64").toString("utf16le");
  assert.equal(decoded, sample);
});

test("Node delta updater script replaces files under Chinese paths and writes applied notice", async () => {
  const root = await fsp.mkdtemp(path.join(os.tmpdir(), "qo-upd-测试中文-"));
  try {
    const resourcesDir = path.join(root, "应用目录", "resources");
    const updatesDir = path.join(root, "用户数据", "updates");
    const stagedDir = path.join(updatesDir, "staged-0.2.23");
    const backupDir = path.join(updatesDir, "backup-0.2.23");
    await fsp.mkdir(path.join(resourcesDir, "qo-runtime", "harness"), { recursive: true });
    await fsp.mkdir(path.join(stagedDir, "qo-runtime", "harness"), { recursive: true });
    await fsp.mkdir(updatesDir, { recursive: true });

    await fsp.writeFile(path.join(resourcesDir, "app.asar"), "old-asar-v0221", "utf8");
    await fsp.writeFile(path.join(resourcesDir, "qo-runtime", "harness", "office.patch.yml"), "old-yml", "utf8");
    await fsp.writeFile(path.join(stagedDir, "app.asar.payload"), "new-asar-v0223", "utf8");
    await fsp.writeFile(path.join(stagedDir, "qo-runtime", "harness", "office.patch.yml.payload"), "new-yml", "utf8");

    const scriptPath = path.join(updatesDir, "apply-delta-0.2.23.cjs");
    const configPath = path.join(updatesDir, "apply-delta-0.2.23.json");
    const appliedPath = path.join(updatesDir, "last-applied.json");
    const errorLogPath = path.join(updatesDir, "last-error.log");
    const logPath = path.join(updatesDir, "last-update.log");
    const startedFlagPath = path.join(updatesDir, "started-0.2.23.flag");

    await fsp.writeFile(scriptPath, buildDeltaNodeScript(), "utf8");
    await fsp.writeFile(configPath, JSON.stringify({
      pid: 99999999,
      resourcesDir,
      stagedDir,
      backupDir,
      exePath: path.join(root, "missing-no-gui-launch.exe"),
      appliedPath,
      errorLogPath,
      logPath,
      startedFlagPath,
      patchFile: "",
      files: ["app.asar", "qo-runtime/harness/office.patch.yml"],
      version: "0.2.23"
    }), "utf8");

    await new Promise((resolve, reject) => {
      const child = spawn(process.execPath, [scriptPath, configPath], {
        cwd: updatesDir,
        windowsHide: true,
        stdio: "ignore"
      });
      child.on("error", reject);
      child.on("exit", (code) => {
        if (code === 0) resolve();
        else reject(new Error(`updater exited with code ${code}`));
      });
    });

    assert.equal(await fsp.readFile(path.join(resourcesDir, "app.asar"), "utf8"), "new-asar-v0223");
    assert.equal(
      await fsp.readFile(path.join(resourcesDir, "qo-runtime", "harness", "office.patch.yml"), "utf8"),
      "new-yml"
    );
    const applied = JSON.parse(await fsp.readFile(appliedPath, "utf8"));
    assert.equal(applied.version, "0.2.23");
    assert.equal(applied.mode, "delta");
    assert.equal(fs.existsSync(errorLogPath), false);
    const logText = await fsp.readFile(logPath, "utf8");
    assert.match(logText, /delta update applied successfully/);
  } finally {
    await fsp.rm(root, { recursive: true, force: true });
  }
});
