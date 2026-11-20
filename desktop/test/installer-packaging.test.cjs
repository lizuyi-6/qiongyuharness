const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const DESKTOP_ROOT = path.resolve(__dirname, "..");

test("运行时裁剪只匹配指定的开发期文件后缀", async () => {
  const { isDevOnlyRuntimeFile } = await import(pathToFileURL(path.join(DESKTOP_ROOT, "scripts", "runtime-prune.mjs")));
  for (const fileName of ["a.d.ts", "a.ts", "a.mts", "a.cts", "a.d.mts", "a.d.cts", "a.js.map", "a.map"]) {
    assert.equal(isDevOnlyRuntimeFile(fileName), true, `${fileName} 应当裁剪`);
  }
  for (const fileName of ["a.js", "a.cjs", "a.mjs", "a.json", "a.node", "a.yml", "a.wasm", "a.md", "ts", "a.tsx.js"]) {
    assert.equal(isDevOnlyRuntimeFile(fileName), false, `${fileName} 应当保留`);
  }
});

test("electron-builder 配置指向 7z 运行时并保留 NSIS 默认压缩类型", () => {
  const packageJson = JSON.parse(fs.readFileSync(path.join(DESKTOP_ROOT, "package.json"), "utf8"));
  assert.equal(packageJson.version, "0.2.4");
  assert.deepEqual(packageJson.build.extraResources, [{ from: "runtime-stage/qo-runtime.7z", to: "qo-runtime.7z" }]);
  assert.deepEqual(packageJson.build.electronLanguages, ["zh-CN", "en-US"]);
  assert.equal(packageJson.build.nsis.oneClick, false);
  assert.equal(packageJson.build.nsis.perMachine, false);
  assert.equal(packageJson.build.nsis.allowToChangeInstallationDirectory, true);
  assert.equal(packageJson.build.nsis.deleteAppDataOnUninstall, true);
  assert.equal(packageJson.build.nsis.include, "installer.nsh");
  assert.deepEqual(packageJson.build.nsis.installerLanguages, ["zh_CN"]);
  assert.equal(packageJson.build.nsis.language, "2052");
  // electron-builder 在 multiLanguageInstaller=false 时会忽略 installerLanguages、强制 en_US(按钮/标题变英文)；
  // 这里直接用它自己的 LangConfigurator 计算实际写进安装器的语言，并确认不弹语言选择框。
  const { LangConfigurator } = require("app-builder-lib/out/targets/nsis/nsisLang.js");
  assert.deepEqual(new LangConfigurator(packageJson.build.nsis).langs, ["zh_CN"]);
  assert.notEqual(packageJson.build.nsis.displayLanguageSelector, true);
  assert.equal(packageJson.build.nsis.installerSidebar, "installerSidebar.bmp");
  assert.equal(packageJson.build.nsis.uninstallerSidebar, "installerSidebar.bmp");
  assert.equal(packageJson.build.nsis.installerHeader, "installerHeader.bmp");
  assert.equal(packageJson.build.nsis.shortcutName, "穹驭办公");
  const compressedExtensions = packageJson.build.nsis.preCompressedFileExtensions;
  const defaults = [".avi", ".mov", ".m4v", ".mp4", ".m4p", ".qt", ".mkv", ".webm", ".vmdk"];
  assert.ok(compressedExtensions.includes(".7z"));
  for (const extension of defaults) assert.ok(compressedExtensions.includes(extension), `${extension} 默认项缺失`);
});

test("NSIS 安装宏直接解压运行时并在校验失败时退出", () => {
  const installer = fs.readFileSync(path.join(DESKTOP_ROOT, "build", "installer.nsh"), "utf8");
  assert.match(installer, /Nsis7z::ExtractWithDetails/);
  assert.match(installer, /nsExec::Exec/);
  assert.doesNotMatch(installer, /ExecWait/);
  assert.doesNotMatch(installer, /deepseek|dsh/i);
  const customInstall = installer.match(/!macro customInstall\n([\s\S]*?)!macroend/);
  assert.ok(customInstall, "缺少 customInstall 宏");
  assert.match(customInstall[1], /qoRuntimeInstallFailed:[\s\S]*?SetErrorLevel 2[\s\S]*?Quit/);
  assert.match(installer, /StrCpy \$isForceCurrentInstall "1"/);
});

function readBmp(fileName, expectedWidth, expectedHeight) {
  const bitmap = fs.readFileSync(path.join(DESKTOP_ROOT, "build", fileName));
  assert.equal(bitmap.toString("ascii", 0, 2), "BM");
  assert.equal(bitmap.readUInt32LE(2), bitmap.length);
  assert.equal(bitmap.readUInt32LE(14), 40);
  assert.equal(bitmap.readInt32LE(18), expectedWidth);
  assert.equal(bitmap.readInt32LE(22), expectedHeight);
  assert.equal(bitmap.readUInt16LE(26), 1);
  assert.equal(bitmap.readUInt16LE(28), 24);
  assert.equal(bitmap.readUInt32LE(30), 0);
  const paddedRowBytes = Math.ceil((expectedWidth * 3) / 4) * 4;
  assert.equal(bitmap.length, 54 + paddedRowBytes * expectedHeight);
}

test("安装器侧栏和页眉为指定尺寸的无压缩 24 位 BMP", () => {
  readBmp("installerSidebar.bmp", 164, 314);
  readBmp("installerHeader.bmp", 150, 57);
});

test("exe 写入穹驭图标与版本信息，窗口和任务栏使用同一图标", () => {
  const packageJson = JSON.parse(fs.readFileSync(path.join(DESKTOP_ROOT, "package.json"), "utf8"));
  // signAndEditExecutable=false 会让 electron-builder 跳过改写 exe 资源，exe 保留 Electron 默认图标(任务栏/快捷方式都显示原子图标)。
  assert.equal(packageJson.build.win.signAndEditExecutable, true);
  assert.equal(packageJson.build.win.icon, "office.ico");
  assert.ok(packageJson.build.files.includes("office.ico"), "office.ico 必须打进 app.asar，窗口图标才能加载");
  const icon = fs.readFileSync(path.join(DESKTOP_ROOT, "office.ico"));
  assert.equal(icon.readUInt16LE(2), 1, "office.ico 必须是 ICO");
  const sizes = Array.from({ length: icon.readUInt16LE(4) }, (_, index) => icon[6 + 16 * index] || 256);
  assert.ok(sizes.includes(256) && sizes.includes(32) && sizes.includes(16), `office.ico 尺寸不全: ${sizes}`);
  const main = fs.readFileSync(path.join(DESKTOP_ROOT, "main.cjs"), "utf8");
  assert.match(main, /icon: path\.join\(__dirname, "office\.ico"\)/);
  assert.ok(main.includes(`app.setAppUserModelId("${packageJson.build.appId}")`), "AppUserModelID 必须等于 build.appId");
});
