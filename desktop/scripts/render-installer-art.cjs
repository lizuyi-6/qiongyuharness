const { app, BrowserWindow } = require("electron");
const fs = require("node:fs/promises");
const path = require("node:path");

app.commandLine.appendSwitch("force-device-scale-factor", "1");
// 图片会被缩放、显示在不同子像素排列的屏幕上：用灰度抗锯齿，避免把 ClearType 彩边烘进位图。
app.commandLine.appendSwitch("disable-lcd-text");
// 依次渲染两张图：第一个窗口销毁后不能让应用按默认行为退出，否则第二张加载会 ERR_FAILED。
app.on("window-all-closed", () => {});

const scriptDirectory = __dirname;
const desktopDirectory = path.resolve(scriptDirectory, "..");
const buildDirectory = path.join(desktopDirectory, "build");
const logoPath = path.resolve(desktopDirectory, "..", "server", "static", "logo.svg");

function bmp24FromBgra(bitmap, width, height) {
  const sourceStride = width * 4;
  if (bitmap.length !== sourceStride * height) {
    throw new Error(`捕获位图字节数异常：${bitmap.length}，预期 ${sourceStride * height}`);
  }

  const rowStride = Math.ceil((width * 3) / 4) * 4;
  const pixelBytes = rowStride * height;
  const bmp = Buffer.alloc(54 + pixelBytes);
  bmp.write("BM", 0, "ascii");
  bmp.writeUInt32LE(bmp.length, 2);
  bmp.writeUInt32LE(54, 10);
  bmp.writeUInt32LE(40, 14);
  bmp.writeInt32LE(width, 18);
  bmp.writeInt32LE(height, 22);
  bmp.writeUInt16LE(1, 26);
  bmp.writeUInt16LE(24, 28);
  bmp.writeUInt32LE(0, 30);
  bmp.writeUInt32LE(pixelBytes, 34);
  bmp.writeInt32LE(2835, 38);
  bmp.writeInt32LE(2835, 42);

  for (let outputRow = 0; outputRow < height; outputRow += 1) {
    const sourceRow = height - 1 - outputRow;
    const sourceOffset = sourceRow * sourceStride;
    const outputOffset = 54 + outputRow * rowStride;
    for (let x = 0; x < width; x += 1) {
      const sourcePixel = sourceOffset + x * 4;
      const outputPixel = outputOffset + x * 3;
      // Windows 上 NativeImage.toBitmap() 是 BGRA，与 BMP 的 BGR 顺序一致，直接取前三个字节。
      bmp[outputPixel] = bitmap[sourcePixel];
      bmp[outputPixel + 1] = bitmap[sourcePixel + 1];
      bmp[outputPixel + 2] = bitmap[sourcePixel + 2];
    }
  }
  return bmp;
}

function imageDataUrl(svg) {
  return `data:image/svg+xml;base64,${Buffer.from(svg, "utf8").toString("base64")}`;
}

function pageHtml({ width, height, background, content }) {
  return `<!doctype html>
<html lang="zh-CN">
<head><meta charset="utf-8"><style>
* { box-sizing: border-box; }
html, body { width: ${width}px; height: ${height}px; margin: 0; overflow: hidden; }
body { background: ${background}; font-family: "Microsoft YaHei UI", "Microsoft YaHei", sans-serif; }
${content}
</style></head>
<body></body>
</html>`;
}

async function captureBmp({ width, height, background, html, output }) {
  // Windows 窗口有最小宽度(约 180px)，开大一点的窗口，再只截取左上角目标区域。
  const window = new BrowserWindow({
    width: Math.max(width, 480),
    height: Math.max(height, 240),
    useContentSize: true,
    show: false,
    backgroundColor: background,
    webPreferences: {
      backgroundThrottling: false,
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  try {
    await window.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    const nativeImage = await window.webContents.capturePage({ x: 0, y: 0, width, height });
    const actualSize = nativeImage.getSize();
    if (actualSize.width !== width || actualSize.height !== height) {
      throw new Error(`${path.basename(output)} 尺寸错误：${actualSize.width}×${actualSize.height}，预期 ${width}×${height}`);
    }
    const bmp = bmp24FromBgra(nativeImage.toBitmap(), width, height);
    await fs.writeFile(output, bmp);
  } finally {
    window.destroy();
  }
}

app.whenReady().then(async () => {
  try {
    const sourceLogo = await fs.readFile(logoPath, "utf8");
    const sidebarLogo = sourceLogo.replace('fill="#18181b"', 'fill="#0b0b0c"');
    const sidebarHtml = pageHtml({
      width: 164,
      height: 314,
      background: "#0b0b0c",
      content: `
        .sidebar { width: 164px; height: 314px; color: #fff; display: flex; flex-direction: column; align-items: center; }
        .brand { margin-top: 70px; display: flex; flex-direction: column; align-items: center; }
        .logo { width: 52px; height: 52px; display: block; }
        .name { margin-top: 12px; font-size: 22px; font-weight: 700; letter-spacing: 1px; }
        .tagline { margin-top: 5px; color: #a1a1aa; font-size: 12px; }
        .footer { margin-top: auto; width: 100%; padding: 0 14px 18px; color: #52525b; text-align: center; font-size: 10px; }
        .rule { height: 1px; margin-bottom: 10px; background: #27272a; }
      `
    });
    const sidebarLogoUrl = imageDataUrl(sidebarLogo);
    const sidebarHtmlWithContent = sidebarHtml.replace("<body></body>", `<body><main class="sidebar"><section class="brand"><img class="logo" src="${sidebarLogoUrl}"><div class="name">穹驭办公</div><div class="tagline">企业办公智能体</div></section><footer class="footer"><div class="rule"></div>qiongyutech</footer></main></body>`);

    const headerHtml = pageHtml({
      width: 150,
      height: 57,
      background: "#ffffff",
      content: `
        .header { width: 150px; height: 57px; padding: 0 10px; display: flex; align-items: center; justify-content: center; gap: 8px; color: #09090b; }
        .logo { width: 26px; height: 26px; display: block; }
        .name { font-size: 14px; font-weight: 700; white-space: nowrap; }
      `
    });
    const headerLogoUrl = imageDataUrl(sourceLogo);
    const headerHtmlWithContent = headerHtml.replace("<body></body>", `<body><header class="header"><img class="logo" src="${headerLogoUrl}"><div class="name">穹驭办公</div></header></body>`);

    await fs.mkdir(buildDirectory, { recursive: true });
    await captureBmp({
      width: 164,
      height: 314,
      background: "#0b0b0c",
      html: sidebarHtmlWithContent,
      output: path.join(buildDirectory, "installerSidebar.bmp")
    });
    await captureBmp({
      width: 150,
      height: 57,
      background: "#ffffff",
      html: headerHtmlWithContent,
      output: path.join(buildDirectory, "installerHeader.bmp")
    });
    console.log("安装界面图片已生成：installerSidebar.bmp（164×314）、installerHeader.bmp（150×57）");
  } catch (error) {
    console.error(error);
    process.exitCode = 1;
  } finally {
    app.quit();
  }
});
