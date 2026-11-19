# 穹驭办公

穹驭办公包含 FastAPI 企业服务、Windows 桌面客户端和本地演示模式。企业部署默认使用 `QO_ENGINE=desktop`：服务端负责身份验证、资料准备和结果保存；员工桌面客户端运行随包提供的 Harness，并通过该员工的 sub2api 密钥直连模型服务。服务端不代员工调用模型。

## 本地演示

Windows 安装 Python 3.12 后运行 `Start-Demo.cmd`；Linux 或 macOS 运行 `./start-demo.sh`。脚本以 `QO_DEMO=1`、`QO_ENGINE=demo` 启动本机演示服务，不需要造场账号或模型密钥。浏览器访问 <http://127.0.0.1:8765>，演示环境提供身份切换入口。演示数据写入 `data/demo`。

## 企业部署

### 1. 配置服务

复制 `.env.example` 为 `.env`，把 `users.example.json` 复制为 `users.json` 并按员工实际信息修改名单。Compose 将名单只读挂载到 `/config/users.json`。填写 PostgreSQL 密码、造场应用凭证和 sub2api 配置；不要将真实密钥提交到版本库。

`.env.example` 示例使用 `QO_ENGINE=desktop`。在此模式下，`QO_GATEWAY_*` 不参与配置校验，也不需要填写。`gateway` 或 `dsh` 引擎才要求 `QO_GATEWAY_BASE_URL`、`QO_GATEWAY_KEY`、`QO_GATEWAY_MODEL`。`QO_DSH_PLUGIN_PATCH` 和 `QO_DSH_PLUGIN_TOOLS` 仅供服务端 `dsh` 引擎使用。

### 2. 登记造场 OIDC 应用

在造场 <https://aetherstudio.top/developers> 登记应用；也可以使用造场的动态注册接口 `POST /api/oauth/register`。应用 issuer 必须是 `https://aetherstudio.top`。登记以下设置：

- 回调地址：`https://office.aethernettech.dev/auth/callback`，按完整地址精确匹配。
- Scope：`openid profile email`。
- Authorization Code 流程启用 PKCE，方法选 `S256`。
- 将应用的 client ID 和 client secret 分别填入 `QO_OIDC_CLIENT_ID`、`QO_OIDC_CLIENT_SECRET`。

服务器校验 OIDC 签名、issuer、audience、有效期、nonce 和需要时的 `azp`。桌面登录经系统浏览器完成，再由本机回环地址接收一次性授权码。工作台会话有效期为 8 小时。

部署时应填入登记应用获得的 client secret。当前启动校验并不强制 secret 非空；只有 secret 非空时，回调才按造场元数据选择 `client_secret_basic` 或 `client_secret_post` 发送。若留空，造场 token endpoint 必须允许无客户端密钥的请求。

### 3. 按邮箱登记员工

`QO_USERS_FILE` 指向 JSON 数组，每项格式为 `{ "email", "name", "department", "role" }`，可选 `sub` 用于预先绑定造场 subject。邮箱会去空格并转成小写；内部用户 ID 由邮箱的 SHA-256 派生，因此同一邮箱在重启后保持稳定。名单至少要有一名 `admin`。示例见 [`users.example.json`](users.example.json)。

登录时先按已绑定的 `sub` 查找启用账号；找不到时，只有造场声明 `email_verified: true` 且名单中存在未绑定的同邮箱账号，才会绑定 subject、写入 `auth.subject_bound` 审计记录并建立会话。邮箱未验证、邮箱不在名单、账号未启用或邮箱已绑定另一个 subject 时，不能通过邮箱路径开通。已经按 subject 命中的账号走优先路径，不会再次检查邮箱验证声明。

从名单删除账号后，**下次服务启动同步名单时**会禁用该账号并删除其会话及未领取桌面授权码。历史业务记录仍保留。

**改动已经绑定 subject 的员工邮箱时，需要管理员清理旧绑定。** 新邮箱会派生出新内部 ID；旧的禁用用户行会保留原 subject，该 subject 唯一索引仍占用它。先在 `users.json` 中更新邮箱并重启服务，再通过受控数据库维护清理旧绑定。先确认旧邮箱行已禁用并记录其内部 ID，再仅将该禁用行的 `subject` 设为 `NULL`；不要删除用户行，以保留历史关联。SQLite 和 PostgreSQL 都可用以下语句（替换旧邮箱）：

```sql
SELECT id, email, enabled, subject FROM users WHERE email = 'old@example.com';
UPDATE users SET subject = NULL WHERE email = 'old@example.com' AND enabled = 0;
```

之后员工用新邮箱登录，系统会把同一造场 subject 绑定到新用户行。若没有清理旧 subject，绑定会被拒绝。

### 4. 配置 sub2api 并开通模型

企业模式要求以下配置：

- `QO_SUB2API_URL`：HTTPS origin，不带 `/v1` 或其他路径，例如 `https://aethernettech.dev`。
- `QO_SUB2API_PROVISION_KEY`：管理员创建的受限管理密钥，只授予 `users.provision.write`，使用 `adm_sc_` 前缀。
- `QO_SUB2API_DEFAULT_GROUP_ID`：正整数分组 ID。
- `QO_SUB2API_KEY_NAME`：密钥名，默认 `qiongyu-office`。

员工登录后，桌面端向工作台调用 `POST /api/v1/me/model-credential`。工作台按员工邮箱和姓名请求 sub2api 找到或创建用户、找到或创建同名密钥，并把密钥明文只返回给桌面端；工作台不把明文密钥写入数据库或审计记录。sub2api 管理员账号所用邮箱不能作为员工开通邮箱：上游返回 `admin_user_not_provisionable` 时，工作台以 403 拒绝。

只有 `employee`、`manager`、`admin` 角色能请求模型凭证；`reviewer` 可以登录工作台，但请求模型凭证会被拒绝。

部署步骤：在 sub2api 中创建只具备上述开通权限的受限管理密钥；确认目标分组 ID；将 URL、密钥、分组 ID 和密钥名填入 `.env`。启动服务后，给员工分发桌面安装包并配置工作台地址。`QO_SUB2API_URL` 不要包含 `/v1`，桌面端会用返回的 `base_url` 请求模型列表及 Chat Completions。

当前服务启动校验检查 sub2api URL 是无路径的 HTTPS origin、provision key 非空、分组 ID 为正整数、key name 符合格式；它不会检查密钥是否以 `adm_sc_` 开头或是否真的只有 `users.provision.write` 权限。密钥的前缀与权限需要在 sub2api 管理端正确设置。

### 5. 启动与安装桌面端

配置 `.env`、`users.json` 和 PostgreSQL 密码后，可用 `docker compose up -d --build` 启动服务。将 `QO_PUBLIC_ORIGIN` 指向的 HTTPS 反向代理转发到服务端 8765 端口。安装包构建命令、运行时目录和本机隔离方式见 [`desktop/README.md`](desktop/README.md)。

桌面模式下，网页端不直接执行 AI 请求；服务端的准备接口会返回本轮授权的提示和上下文，桌面执行后再将文本、模型名和客户端上报的 usage 交回服务端保存。非桌面浏览器发起执行时会得到“请在穹驭办公桌面客户端中使用 AI 办理”的提示。

## 管理员模型权限

- **给员工换模型：**在 sub2api 后台修改该员工 `qiongyu-office` 密钥所属分组。桌面端每次打开会话以及每次执行前都会重新读取该密钥的 `/v1/models` 列表。
- **收回模型权限：**在 sub2api 停用该密钥或该用户，或将员工从 QO 名单移除。按 sub2api 开通接口契约，已停用、过期或额度用尽的同名密钥返回 403 `key_disabled`，不会被自动补发；用户停用返回 403 `user_disabled`。删除密钥会被视为重置；下次登录请求开通时，接口会重新创建同名密钥。名单移除则在下一次服务启动同步时禁用工作台账号并清会话。
- **管理员邮箱：**不要用 sub2api 管理员账号绑定的邮箱作为员工名单邮箱，否则开通接口会返回 403 `admin_user_not_provisionable`。

## 已知限制

- 中国大陆 IP 直连 sub2api 会返回 451；桌面端通过 Electron `net.fetch` 请求模型服务并使用 Electron 系统代理配置。模型列表路径会把 451 显示为“当前网络无法访问模型服务(451)”；生成请求经本机回环网关转发，非 2xx 上游响应会被网关改成 502。工作台服务端访问 OIDC 和 sub2api 的请求使用 `httpx` 且设置 `trust_env=False`，不会继承环境代理。
- 如果 sub2api 分组中的上游账号开启了“账号级破甲”，sub2api 会将 `max_tokens` 提高到 8192 并改写提示词，可能破坏穹驭办公要求的 JSON 输出格式。给穹驭办公使用的分组应选择未开启该选项的账号。
- `runs.usage` 中保存的模型名和用量来自桌面客户端上报；计费与实际用量以 sub2api 为准。
- 桌面本机回环网关将每轮 `max_tokens` 限制在 3500 以内，每轮最多转发 4 次模型调用，只放行本轮授权的 `office_context_read` 工具；未启用上下文工具的请求不放行工具调用。

## 验证命令

在仓库根目录运行：

```text
python -m unittest discover -s tests -v
PYTHONPATH=. python tests/dsh_integration.py
node --test desktop/test
node desktop/scripts/check-no-dsh-conflict.mjs --config-only
```

Windows PowerShell 可先运行 `$env:PYTHONPATH = "."`，再执行 DSH 集成命令。`--config-only` 只检查打包配置；不检查 `desktop/dist/win-unpacked` 中实际生成的安装目录。完整 `npm run build:win` 会在打包后运行冲突检查脚本。
