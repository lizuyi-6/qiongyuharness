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

### 3. 配置名单与开放注册

`QO_USERS_FILE` 指向 JSON 数组，每项格式为 `{ "email", "name", "department", "role" }`，可选 `sub` 用于预先绑定造场 subject，也可用布尔值 `enabled` 停用某个名单账号。邮箱会去空格并转成小写；内部用户 ID 由邮箱的 SHA-256 派生，因此同一邮箱在重启后保持稳定。企业模式名单仍至少要有一名 `admin`。示例见 [`users.example.json`](users.example.json)。

`QO_OPEN_ENROLLMENT` 只接受 `0` 或 `1`，缺省为 `0`。值为 `0` 时维持仅名单登录；同步会停用不在名单里的用户并删除其会话及桌面授权码。生产开放注册时将它设为 `1`：造场会员即可登录；其 ID token 必须通过现有签名、issuer、audience、nonce 和有效期校验，并且新用户的 `email_verified` 必须严格为布尔值 `true`。名单外新用户会以 `employee` 角色创建，姓名取 ID token 的 `name`，缺失时使用邮箱前缀，部门为空；自动开户事件 `user.auto_enrolled` 只记录用户 ID 和邮箱 SHA-256，不记录 token。已绑定的邮箱仍不能换绑到另一个造场 subject。

开放注册模式下，名单仍决定名单内账号的角色、部门和姓名；未列入名单的自动注册账号不会被名单同步停用。将条目设为 `"enabled": false` 会拒绝该邮箱登录，并在下次服务启动同步名单时删除其现有会话和未领取桌面授权码，例如：

```json
{ "email": "former@example.com", "name": "停用账号", "department": "运营", "role": "employee", "enabled": false }
```

造场账号停用或取消会员资格后，造场会撤销其登录凭证；需要在穹驭办公立即封禁的账号，也应在名单中显式设为 `enabled: false`。开放注册不会自动发放模型额度：新用户开通后 sub2api 余额为 0，需管理员在 sub2api 分配余额、订阅或额度后才能对话。

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

部署步骤：在 sub2api 中创建只具备上述开通权限的受限管理密钥；确认目标分组 ID；将 URL、密钥、分组 ID 和密钥名填入 `.env`。部署前在同一 shell 加载这两个环境变量并运行 `python scripts/check_sub2api_key.py`；探测不通过时不要启动服务。启动服务后，给员工分发桌面安装包并配置工作台地址。`QO_SUB2API_URL` 不要包含 `/v1`，桌面端会用返回的 `base_url` 请求模型列表及 Chat Completions。

服务启动校验会检查 sub2api URL 是无路径的 HTTPS origin、provision key 去除首尾空白后以 `adm_sc_` 开头且带有后缀、分组 ID 为正整数、key name 符合格式。它不会向 sub2api 发请求；部署探测脚本会用无效邮箱验证开通接口返回 `400 invalid_email`，并验证读取用户列表返回 `403`。这些探测不替代在 sub2api 管理端确认密钥权限配置。

### 5. 启动与安装桌面端

配置 `.env`、`users.json` 和 PostgreSQL 密码后，可用 `docker compose up -d --build` 启动服务。将 `QO_PUBLIC_ORIGIN` 指向的 HTTPS 反向代理转发到服务端 8765 端口。安装包构建命令、运行时目录和本机隔离方式见 [`desktop/README.md`](desktop/README.md)。

桌面模式下，网页端不直接执行 AI 请求；服务端的准备接口会返回本轮授权的提示和上下文，桌面执行后再将文本、模型名和客户端上报的 usage 交回服务端保存。非桌面浏览器发起执行时会得到“请在穹驭办公桌面客户端中使用 AI 办理”的提示。

## 管理员模型权限

- **给员工换模型：**在 sub2api 后台修改该员工 `qiongyu-office` 密钥所属分组。桌面端每次打开会话以及每次执行前都会重新读取该密钥的 `/v1/models` 列表。
- **收回模型权限：**在 sub2api 停用该密钥或该用户，或将员工从 QO 名单移除。按 sub2api 开通接口契约，已停用、过期或额度用尽的同名密钥返回 403 `key_disabled`，不会被自动补发；用户停用返回 403 `user_disabled`。删除密钥会被视为重置；下次登录请求开通时，接口会重新创建同名密钥。名单移除则在下一次服务启动同步时禁用工作台账号并清会话。
- **管理员邮箱：**不要用 sub2api 管理员账号绑定的邮箱作为员工名单邮箱，否则开通接口会返回 403 `admin_user_not_provisionable`。

## 桌面智能分工

桌面聊天模型列表顶部提供“智能分工（理解→执行→验收）”。角色模型和阶段提示词由已登录桌面客户端从 `/api/v1/me/pipeline` 获取；配置失败或理解、执行、验收任一模型未包含在员工实时 `/v1/models` 列表中时，智能分工不可用，普通模型仍可选择。轻量和复核角色缺少模型时分别按标准流程执行、或将复核记为不可用并保留验收结论。

服务端环境变量：

- `QO_PIPELINE_ENABLED` 只接受 `0` 或 `1`，缺省为 `1`。
- `QO_PIPELINE_MODEL_UNDERSTAND` 缺省 `glm-5.3`。
- `QO_PIPELINE_MODEL_EXECUTE` 缺省 `step-5-preview`。
- `QO_PIPELINE_MODEL_ACCEPT` 缺省 `glm-5.3`。
- `QO_PIPELINE_MODEL_REVIEW` 缺省 `k3.1-pro`。
- `QO_PIPELINE_MODEL_LIGHT` 缺省 `deepseek-v4.1-flash-1017`。

模型值须为 1–200 个字符且不含空白或控制字符；无效配置会阻止服务启动。上述缺省模型 ID 于 2026-11-19 与 sub2api 穹驭分组（id 22）的 `/v1/models` 列表逐一比对过，五个 ID 均在列表中，并经 `/v1/chat/completions` 实测返回 200；分组模型变动后应重新核对并按需覆盖。流水线最多执行两轮，阶段调用上限为 8 次，全程时限为 270 秒；本机按每阶段最多 95 秒执行。流水线遇到额度 429 后会阻止该阶段的后续上游请求，并停止后续阶段；单模型执行仍遵循原有额度恢复行为。线程记录的验收和复核徽标来自服务端保存的流水线 trace。

## 已知限制

- 中国大陆 IP 直连 sub2api 会返回 451；桌面端通过 Electron `net.fetch` 请求模型服务并使用 Electron 系统代理配置。模型列表路径会把 451 显示为“当前网络无法访问模型服务(451)”；生成请求经本机回环网关转发，非 2xx 上游响应会被网关改成 502。工作台服务端访问 OIDC 和 sub2api 的请求使用 `httpx` 且设置 `trust_env=False`，不会继承环境代理。
- 如果 sub2api 分组中的上游账号开启了“账号级破甲”，sub2api 会将 `max_tokens` 提高到 8192 并改写提示词，可能破坏穹驭办公要求的 JSON 输出格式。给穹驭办公使用的分组应选择未开启该选项的账号。
- `runs.usage` 中保存的模型名和用量来自桌面客户端上报；计费与实际用量以 sub2api 为准。
- 桌面本机回环网关将每轮 `max_tokens` 限制在 3500 以内：非正整数请求会拒绝，超过 3500 的请求会夹逼并在结果旁提示；上游 `finish_reason=length` 时也会提示回复被截断。每轮最多转发 4 次模型调用，只放行本轮授权的 `office_context_read` 工具；未启用上下文工具的请求不放行工具调用。

## 验证命令

在仓库根目录运行：

```text
python -m unittest discover -s tests -v
PYTHONPATH=. python tests/dsh_integration.py
node --test desktop/test
node desktop/scripts/check-no-dsh-conflict.mjs --config-only
```

Windows PowerShell 可先运行 `$env:PYTHONPATH = "."`，再执行 DSH 集成命令。`--config-only` 只检查打包配置；不检查 `desktop/dist/win-unpacked` 中实际生成的安装目录。完整 `npm run build:win` 会在打包后运行冲突检查脚本。
