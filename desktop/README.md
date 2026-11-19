# 穹驭办公桌面客户端

企业桌面模式使用 `QO_ENGINE=desktop`。服务端负责登录、准备本轮输入和保存结果；Electron 主进程在员工电脑上运行安装包内的 Harness，直接连接该员工的 sub2api 模型密钥。桌面端不把模型密钥交给网页渲染进程或 Harness 子进程。

## Windows 构建

在 `desktop` 目录安装锁定的开发依赖后运行：

```text
npm ci
npm run build:win
```

`build:win` 依次运行 `scripts/prepare-runtime.mjs`、Electron Builder 的 Windows x64 打包、`scripts/check-no-dsh-conflict.mjs`。准备脚本将 Harness 的 `package.json`、`package-lock.json`、`office.patch.yml`、`office-context.mjs` 复制到 `runtime-stage/qo-runtime/harness/`，执行 `npm ci --omit=dev --ignore-scripts` 安装运行依赖，并把构建所用且版本至少为 24.9.0 的 Node 可执行文件暂存为 `runtime-stage/qo-runtime/node/node.exe`。构建要求 x64 Node。

Windows 安装器只由 `desktop/package.json` 的 electron-builder `build.nsis` 配置生成；`build.nsis.include` 使用 `desktop/build/uninstaller.nsh`，保留 electron-builder 默认的卸载/升级清理逻辑，并在普通卸载时增加支持长路径的 `$INSTDIR` 清理。升级分支跳过这个额外清理，由 electron-builder 默认升级流程接管。普通卸载会删除 `%APPDATA%\QiongyuOffice`，其中包含连接配置和 safeStorage 加密的模型凭证；覆盖升级保留该目录。

Electron Builder 配置将运行时放到安装目录的 `resources/qo-runtime/`。Electron `files` 只收主进程、桌面模块、页面和 `package.json`，不收 `harness` 或 `runtime-stage`；按此配置 Harness 位于 `resources/app.asar`、`resources/app/` 和 `resources/app.asar.unpacked/` 之外。桌面端从 `resources/qo-runtime/` 解析 Harness 和 Node 路径；打包后的实际目录仍应由构建末尾的检查脚本核对。

打包完成后冲突检查脚本扫描 `dist/win-unpacked/resources`，检查 Harness 位于 `resources/qo-runtime/`、归档内没有冲突的 `node_modules/@deepseek-ai/` 路径，以及产品名、安装器和隔离设置。单独运行 `node scripts/check-no-dsh-conflict.mjs --config-only` 只检查配置，不会扫描已生成的安装目录。

## 登录、密钥和模型

- 桌面登录在 `127.0.0.1` 随机端口启动一次性回调监听，浏览器登录等待上限为 120 秒。服务端验证造场回调后，将短时一次性授权码重定向到本机；桌面端用 PKCE verifier 兑换工作台会话。授权码有效期为 2 分钟且只能领取一次。
- 登录后主进程预取模型凭证。请求使用窗口的 Electron session cookie 调用工作台 `/api/v1/me/model-credential`。主进程用 Electron `safeStorage` 加密密钥，并按服务 origin 和内部用户 ID 分开存放于 `%APPDATA%\QiongyuOffice\qo-model-credentials\`。切换服务地址或退出登录时删除相应凭证；系统不可用 `safeStorage` 时不保存明文替代。
- 每次打开会话时从该密钥的 sub2api `/v1/models` 读取可用模型；执行前再读取一次。若员工上次选择的模型已不在列表中，选择会被清空并提示重新选择。每次运行通过 Electron 网络栈访问 sub2api，使用系统代理。
- 工作台会话和 sub2api 密钥只供主进程使用。桌面端把选定模型与提示交给本机 Harness，经本机回环网关转发；Harness 子进程只拿到临时回环 bearer key，不拿到 sub2api 用户密钥。

## 每轮执行的隔离与限制

每轮执行都会创建临时目录，并将子进程的 `HOME`、`USERPROFILE`、`DSH_HOME`、`DSH_AGENTS_HOME` 和工作目录指向该目录。结束或报错后，主进程会结束子进程、关闭回环网关并删除临时目录。子进程继承的系统环境变量限于 `PATH`、`SystemRoot`、`WINDIR`、`TEMP`、`TMP`、`LANG`，另加本轮运行所需变量及 `NO_PROXY=127.0.0.1,localhost`。

本机网关只接收带本轮随机 bearer key 的 `POST /v1/chat/completions`：

- 请求体最大 2 MiB，响应最大 3 MiB；请求中的模型必须等于本轮选择的模型。
- `max_tokens` 缺省或为 0 时使用 3500，超过 3500 时夹到 3500；小于 1 或非整数输入会按请求格式错误拒绝。
- 每轮最多接受 4 次模型调用；只有本轮上下文授权的 `office_context_read` 工具可以出现，未授权时不接受工具或工具调用。其他远程请求由 DSH 桥拒绝。
- Harness 使用 `sdk-minimal` profile，不使用名为 `desktop` 的 profile。DSH 调用超时为 95 秒。

这是本地请求边界，sub2api 仍是模型授权、分组策略、计费和实际用量的依据。若上游分组启用了“账号级破甲”，其提示词改写可能与穹驭办公的 JSON 输出约定冲突；为穹驭办公分组选择未开启破甲的上游账号。

## 与本机 DeepSeek Harness / dsh-purge 的隔离约定

- Harness 运行时在 `resources/qo-runtime/`，不会纳入 `app.asar`、`app/` 或 `app.asar.unpacked/`。根 `package.json` 名称为 `qiongyu-office-desktop`；`productName` 为“穹驭办公”，exe 名随产品名生成，安装器文件名为 `QiongyuOffice-${version}-${arch}.exe`；快捷方式、安装目录和卸载注册表显示名使用穹驭办公名称，不使用 DeepSeek Harness、DSH Desktop 或 `deepseek-harness` 名称。
- 每轮临时目录隔离 `HOME`、`USERPROFILE`、`DSH_HOME`、`DSH_AGENTS_HOME` 和 cwd；不使用 `~/.dsh` 或 `~/.agents`。不注册全局 `dsh` 命令，也不修改系统 `PATH`。
- 使用 `sdk-minimal` profile。卸载器只删除当前穹驭办公安装目录 `$INSTDIR`、穹驭办公自己的 `%APPDATA%\QiongyuOffice` 用户数据、快捷方式和穹驭办公卸载注册表项；普通卸载删除用户数据，覆盖升级保留用户数据。
- 每次 `build:win` 后运行的冲突检查会核对打包配置和实际 `dist/win-unpacked` 产物；命令带 `--config-only` 时只核对配置。
