# 穹驭办公 0.1.0

穹驭科技企业桌面 AI 办公 Agent。管理后台面向主管和审核人，桌面工作台面向员工。当前交付包含 Electron Windows 客户端、FastAPI 服务、SQLite 演示和 PostgreSQL 配置、DeepSeek Harness 接入和 Word/PPT 文稿处理。

## 本地运行

1. 解压源码。Windows 安装 Python 3.12 后双击 `Start-Demo.cmd`。Linux 或 macOS 执行 `./start-demo.sh`。首次启动需要下载 Python 依赖。启动终端保持打开。
2. 在浏览器打开 [http://127.0.0.1:8765](http://127.0.0.1:8765)，或者安装 Windows x64 客户端，填写相同服务地址。安装器按当前用户安装，桌面客户端可从开始菜单启动。
3. 首页默认进入员工林知夏的办公会话。单击右上角「演示环境」可切换李祖祎（超级管理员）、周婧（部门主管）、其他员工或顾言（审核人）。带 `?login=1` 可先查看身份选择页。
4. 使用超级管理员登录，进入任务管理并创建任务。勾选一名或多名员工，可指定截止时间、审核人和任务材料。保存草稿或发布任务。
5. 切到所选员工，在「办公会话」右侧选定任务，或从任务详情点击「用会话办理」。可以先梳理要求、引用资料，再用自然语言请求生成或修订 Word/PPT。可以在对话内查看执行记录、预览和下载文件。
6. 点击「回到任务交付」，确认成果版本后提交。管理员或指定审核人在审核队列验收或退回。办公导航的「文稿」支持自行编写和导出大纲。

生成 PPT 后，右侧会显示逐页预览。选择一页，或点击预览中的可编辑文本框，再点击「让 AI 改选区」并发送修改要求。修订版仍是可编辑 PPTX，原稿和其他页留在原有版本链中。当前范围是本会话生成的 PPT 与已有单段文本框，复杂母版、字体替换等版式以 PowerPoint 的显示为准。

在「项目」新建项目并把邀请码交给同学。同学使用自己的账号加入后，可以发布会议纪要、要求、决策或共享个人文件。项目会话自动参考最近六条共享记录的可提取正文，各成员的会话记录仍由本人查看。源码部署或本地演示启动时会自动添加项目和会话关联表。

演示模式使用本地规则和文稿模板并显示提示，无须提供网关密钥。演示模式仅面向本机，用于验收流程。演示数据库位于源码目录的 `data/demo`。

## 企业部署

1. 复制 `.env.example` 为 `.env`，将 `users.example.json` 复制为 `users.json`。把名单的 `sub` 替换为该 OIDC 客户端实际颁发的 Subject，按公司部门指定角色。跨部门授权由管理员承担。服务启动时按 `sub` 同步名单。移除账号并重启服务后，其现有会话失效，历史任务仍保留。
2. 在公司身份平台登记 `<QO_PUBLIC_ORIGIN>/auth/callback`。应用使用 Authorization Code + PKCE S256、state 和 nonce，校验 ID Token 签名、颁发者、受众及有效期。桌面客户端通过系统浏览器登录，使用一次性领取凭证建立浏览器容器内的内存会话。本版本的工作台会话有效 8 小时，过期时重新登录。
3. 填写企业网关的 HTTPS Base URL（通常以 `/v1` 结尾）、服务端密钥和模型名称。网关需兼容 OpenAI Chat Completions。在 `dsh` 模式还需返回标准 SSE 流，支持 DeepSeek 使用的流式参数和 usage。如果网关只支持非流式的标准 OpenAI 协议，可设置 `QO_ENGINE=gateway`。生产网关凭证只保存在服务端环境。
4. 为 PostgreSQL 填写随机密码，采用十六进制字符可直接用于 Compose 的数据库连接字符串。执行 `docker compose up -d --build`，用公司的 HTTPS 反向代理把 `QO_PUBLIC_ORIGIN` 转发至本机 8765。容器使用独立数据库卷，升级和恢复时同时备份数据库及配置。可按 `QO_DATABASE_URL` 接入外部 PostgreSQL。
5. 管理员分发 Windows 安装包和企业地址。可通过 `--server=https://office.example.com` 或 `QO_DESKTOP_URL` 预设。设置公司内网扫描服务地址 `QO_CLAMD=host:3310` 后，外部上传文件经过 ClamAV INSTREAM 检查。文件上限 12 MB。
6. 建议把运行容器的出网范围约束到公司身份服务、公司网关和内部扫描服务。网关收到 `X-Request-ID`、`X-QO-Employee`（Subject 的 UTF-8 十六进制）、`X-QO-Assignment` 和 `X-QO-Run`，可用于对账和限额。身份服务使用外部 HTTPS 时，需保证容器信任其 CA。

编译 Windows 安装包的标准方式是在 Windows 上从 `desktop` 执行 `npm ci` 和 `npm run build:win`。Linux 交叉编译可先运行 Electron Builder 的 `--win --x64 --dir`，再从该目录执行 `makensis installer.nsi`。企业正式发放时使用自有代码签名证书签名，并在公司 Windows 镜像上验收。

## 已实现的边界

| 模块 | 交付功能 |
|---|---|
| 任务 | 部门范围、单人或批量派发、草稿、定时发布、按工作日重复及指定例外、处理状态、延期申请、转派、任务历史 |
| 交付 | 文件关联、HTTPS 链接、独立提交版本、验收与退回、历史来源链、访问审计 |
| 会话 | 历史会话、多轮跟进、任务与项目关联、流式执行状态、会话内预览和下载成果 |
| 文稿 | Word/PPT 可编辑导出、PPT 逐页预览、页内选区改稿、修订副本、手工大纲编辑 |
| 项目 | 邀请码、成员访问、共享笔记和文件、项目会话引用 |
| 身份 | OIDC 网页与桌面流、服务端资源鉴权、桌面凭证在内存分区、演示角色 |
| 模型 | 固定 DeepSeek Harness 0.2.2-rc.1、受控公司网关、模型会话和 usage 对账、单次输出上限 |
| 数据 | SQLite 或 PostgreSQL、外键和事务、幂等派发与提交、任务版本条件、轮询与增量 API |

本版的 Harness 和 Office 文稿渲染部署在受控服务端，Electron 负责 GUI、系统浏览器登录和下载。Harness 使用官方 SDK JSON-RPC 加定制配置。Office 操作用本项目的文稿处理模块。Dataelem 桌面底座、Windows COM 本地执行、PDF 转换、企业模板库、签名自动更新及完整离线同步属于后续工作。任务备注和编辑中的大纲在打开的窗口内可以恢复，任务列表在线时约每 4 秒刷新。文件保存在数据库，独立对象存储及保留期管理属于生产化扩展。

应用保存会话历史。每一轮向 Harness 提供近期对话、关联任务要求和所选资料的正文，由服务端执行可编辑 Word/PPT 导出。导出的文件进入个人空间，提交到任务时由员工确认。升级启动时会自动补建会话数据表。

通用文件预览使用正文提取。PPT 会话侧栏使用形状、文字和位图绘制逐页预览。原有的局部替换支持同一段落的连续文本，跨格式段落或特殊对象可用本地 Office/WPS 编辑。XLSX/PDF 可作为材料或交付附件保存，Excel 内容编辑、PDF 原生编辑和扫描件 OCR 属于后续功能。

## DeepSeek Harness 插件

项目会话和选页修改通过原生 Cordis 工具 `office_context_read` 读取本轮授权的上下文。工具使用 `@deepseek-ai/dsh-tools` 注册，模型调用的网关须支持流式工具调用。`QO_ENGINE=gateway` 使用普通请求正文传递相同授权数据，演示模式使用本地规则。

可通过服务器配置 `QO_DSH_PLUGIN_PATCH` 和 `QO_DSH_PLUGIN_TOOLS` 接入原生插件。前者为服务端可读取的 Cordis 补丁路径，后者为允许暴露的函数名列表。插件代码与依赖须部署在服务器的 Harness 运行环境，并在锁定的 `0.2.2-rc.1` 版本验收。例如，可将两项分别设为 `/opt/office/harness/todo.example.patch.yml` 和 `todo_write`，加载上游原包的任务规划工具。此示例默认关闭，计划卡片若需出现在办公界面，需要单独映射原生事件。

纯工具类插件可优先按依赖和版本测试。依赖上游页面、工作目录写入、终端或 Office sidecar 的插件需要逐项适配，企业侧继续使用自己的身份和文件访问范围。服务器只加载审核过的信任代码，网关函数名名单用于约束模型暴露面。

## API 与验证

本地演示的 API schema 为 [http://127.0.0.1:8765/openapi.json](http://127.0.0.1:8765/openapi.json)。主接口包括 `POST /api/v1/tasks`、`POST /api/v1/tasks/{id}/publish`、`GET /api/v1/me/today`、`POST /api/v1/assignments/{id}/sessions`、`POST /api/v1/agent/render`、`POST /api/v1/assignments/{id}/submissions` 和 `POST /api/v1/submissions/{id}/review`。派发、发布、提交、审核写接口接收 `Idempotency-Key`，修改任务和状态变化接收 `version`。

- `GET/POST /api/v1/conversations` 用于浏览和新建会话。`POST /api/v1/conversations/{id}/turns` 以 NDJSON 返回执行进度和结果。
- `GET /api/v1/files/{id}/slides` 与 `GET /api/v1/files/{id}/slides/{page}` 返回页列表和预览数据。`/api/v1/projects` 提供项目成员与共享资料入口。
- 在项目目录运行 `python -m unittest discover -s tests -v`。用例覆盖跨账号数据隔离、文件版本、防重复执行、定时计划、会话生成与修订及 OIDC 桌面凭证的单次领取。
- 使用 Node.js 24 执行 `npm ci --prefix harness`，再执行 `PYTHONPATH=. python tests/dsh_integration.py` 验证实际 DSH 进程到模拟流式网关的请求、Usage 统计和工具限制。Windows PowerShell 使用 `$env:PYTHONPATH = "."` 设置 Python 模块路径。
- Chromium 验收走通了任务派发、引用资料、两轮 PPT 生成与修订、重开会话、提交和验收。Windows 安装包已完成交叉编译，企业账号及真实 Windows 环境实测需部署后进行。
- `harness/package-lock.json` 和 `desktop/package-lock.json` 锁定 npm 依赖。DeepSeek Harness 采用 MIT 许可证，源码位于 [deepseek-ai/deepseek-harness](https://github.com/deepseek-ai/deepseek-harness)。办公文件使用 python-docx 和 python-pptx 原生 Open XML 生成。安装包包含 Electron 授权文件。
