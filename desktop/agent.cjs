const { net, safeStorage } = require("electron");
const { createHash, randomUUID } = require("node:crypto");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { createLoopbackGateway } = require("./lib/loopback-gateway.cjs");
const { runDsh } = require("./lib/dsh-bridge.cjs");
const { listModelsWithRefresh, modelServiceError, parseModels } = require("./lib/models.cjs");
const { buildRuntimeEnv, resolveRuntimePaths } = require("./lib/runtime-env.cjs");
const { sessionFetch } = require("./lib/session-cookie.cjs");

function digest(value) {
  return createHash("sha256").update(value).digest("hex");
}

function nodeAtLeast(version, minimum = [24, 9, 0]) {
  const found = String(version).split(".").map((part) => Number.parseInt(part, 10));
  for (let index = 0; index < minimum.length; index += 1) {
    if ((found[index] || 0) > minimum[index]) return true;
    if ((found[index] || 0) < minimum[index]) return false;
  }
  return true;
}

function validateBaseUrl(value) {
  let url;
  try { url = new URL(value); } catch { throw new Error("模型凭证返回的服务地址无效"); }
  if (url.protocol !== "https:" || url.username || url.password || url.pathname !== "/v1" || url.search || url.hash) {
    throw new Error("模型凭证返回的服务地址无效");
  }
  return url.origin + "/v1";
}

function responseError(payload, status, fallback) {
  if (typeof payload?.detail === "string" && payload.detail.trim()) return payload.detail.trim();
  return modelServiceError(status, undefined) || fallback;
}

function createDesktopAgent({ app, getWindow, getOrigin, emit = () => {} }) {
  const credentialRoot = path.join(app.getPath("userData"), "qo-model-credentials");
  const runtimePaths = resolveRuntimePaths(app.isPackaged ? process.resourcesPath : undefined, __dirname);

  function currentSession() {
    const window = getWindow();
    if (!window || window.isDestroyed()) throw new Error("桌面窗口已关闭");
    return window.webContents.session;
  }

  function credentialDirectory(origin) {
    return path.join(credentialRoot, digest(origin));
  }

  function credentialFile(origin, userId) {
    return path.join(credentialDirectory(origin), `${digest(userId)}.json`);
  }

  async function fetchSessionJson(origin, url, options = {}) {
    const response = await sessionFetch(currentSession(), origin, url, {
      redirect: "error",
      ...options
    });
    const raw = await response.text();
    let payload = {};
    if (raw) {
      try { payload = JSON.parse(raw); } catch {
        if (response.ok) throw new Error("工作台返回了无效数据");
      }
    }
    if (!response.ok) throw new Error(responseError(payload, response.status, "工作台请求失败"));
    return payload;
  }

  async function identityFor(origin) {
    const user = await fetchSessionJson(origin, `${origin}/api/v1/me`);
    if (typeof user.id !== "string" || !user.id) throw new Error("工作台用户信息无效");
    return user;
  }

  async function readCredential(origin, userId) {
    if (!safeStorage.isEncryptionAvailable()) throw new Error("当前系统无法安全保存模型凭证");
    try {
      const raw = await fs.readFile(credentialFile(origin, userId), "utf8");
      const saved = JSON.parse(raw);
      if (saved.origin !== origin || saved.userId !== userId || typeof saved.baseUrl !== "string" ||
          typeof saved.encryptedKey !== "string") return null;
      const apiKey = safeStorage.decryptString(Buffer.from(saved.encryptedKey, "base64"));
      if (!apiKey) return null;
      return { baseUrl: validateBaseUrl(saved.baseUrl), apiKey };
    } catch (error) {
      if (error.code === "ENOENT") return null;
      await fs.rm(credentialFile(origin, userId), { force: true }).catch(() => {});
      return null;
    }
  }

  async function saveCredential(origin, userId, credential) {
    if (!safeStorage.isEncryptionAvailable()) throw new Error("当前系统无法安全保存模型凭证");
    const directory = credentialDirectory(origin);
    await fs.mkdir(directory, { recursive: true });
    const encryptedKey = safeStorage.encryptString(credential.apiKey).toString("base64");
    const file = credentialFile(origin, userId);
    const temporary = `${file}.${randomUUID()}.tmp`;
    await fs.writeFile(temporary, JSON.stringify({
      origin, userId, baseUrl: credential.baseUrl, encryptedKey
    }), { encoding: "utf8", mode: 0o600 });
    await fs.rename(temporary, file);
  }

  async function requestCredential(origin, user) {
    const payload = await fetchSessionJson(origin, `${origin}/api/v1/me/model-credential`, {
      method: "POST"
    });
    if (typeof payload.api_key !== "string" || !payload.api_key ||
        !Number.isSafeInteger(payload.key_id) || !Number.isSafeInteger(payload.group_id)) {
      throw new Error("工作台返回的模型凭证无效");
    }
    const credential = { apiKey: payload.api_key, baseUrl: validateBaseUrl(payload.base_url) };
    await saveCredential(origin, user.id, credential);
    return credential;
  }

  async function getCredential(origin = getOrigin(), { force = false } = {}) {
    if (!origin) throw new Error("请先连接穹驭办公服务");
    const user = await identityFor(origin);
    const stored = force ? null : await readCredential(origin, user.id);
    return { user, credential: stored || await requestCredential(origin, user) };
  }

  async function clearCredentialsForOrigin(origin) {
    if (!origin) return;
    await fs.rm(credentialDirectory(origin), { recursive: true, force: true });
  }

  async function clearAllCredentials() {
    await fs.rm(credentialRoot, { recursive: true, force: true });
  }

  async function fetchModels(credential) {
    const url = `${credential.baseUrl.replace(/\/+$/, "")}/models`;
    const response = await net.fetch(url, {
      headers: { Authorization: `Bearer ${credential.apiKey}`, Accept: "application/json" },
      redirect: "error"
    });
    if (!response.ok) {
      let payload = {};
      try { payload = await response.json(); } catch {}
      throw Object.assign(new Error(modelServiceError(response.status, payload?.detail)), { status: response.status });
    }
    let payload;
    try { payload = await response.json(); } catch { throw new Error("模型服务返回的列表格式无效"); }
    return parseModels(payload);
  }

  async function freshModels() {
    const origin = getOrigin();
    return listModelsWithRefresh(
      async (force) => (await getCredential(origin, { force })).credential,
      fetchModels
    );
  }

  async function models() {
    try {
      return { models: await freshModels() };
    } catch (error) {
      return { models: [], error: error.message || "模型列表加载失败" };
    }
  }

  async function prefetchAfterLogin() {
    const origin = getOrigin();
    await getCredential(origin, { force: true });
    return freshModels();
  }

  function nodeExecutable() {
    if (require("node:fs").existsSync(runtimePaths.bundledNode)) {
      return { executable: runtimePaths.bundledNode, extraEnv: {} };
    }
    if (!nodeAtLeast(process.versions.node)) throw new Error("桌面运行时 Node 版本低于 24.9.0");
    return { executable: process.execPath, extraEnv: { ELECTRON_RUN_AS_NODE: "1" } };
  }

  async function runAgent(job) {
    if (!job || typeof job.system !== "string" || typeof job.prompt !== "string" || typeof job.model !== "string") {
      throw new Error("桌面执行请求无效");
    }
    const origin = getOrigin();
    emit({ type: "step", label: job.context ? "载入项目资料" : "准备执行环境" });
    const available = await freshModels();
    if (!available.includes(job.model)) throw new Error("该模型已不再对你开放，请重新选择");
    const { credential } = await getCredential(origin);

    const place = await fs.mkdtemp(path.join(os.tmpdir(), "qo-office-agent-"));
    let gateway;
    try {
      let contextPath;
      let contextPatchPath;
      if (job.context !== undefined && job.context !== null) {
        contextPath = path.join(place, "office-context.json");
        await fs.writeFile(contextPath, JSON.stringify(job.context), { encoding: "utf8", mode: 0o600 });
        contextPatchPath = path.join(place, "office-context.patch.yml");
        await fs.writeFile(contextPatchPath,
          `- insert:\n    - id: office-context\n      name: ${JSON.stringify(runtimePaths.contextPlugin)}\n`,
          { encoding: "utf8", mode: 0o600 });
      }

      gateway = await createLoopbackGateway({
        baseUrl: credential.baseUrl,
        apiKey: credential.apiKey,
        model: job.model,
        allowedTools: contextPath ? ["office_context_read"] : [],
        fetchImpl: (url, options) => net.fetch(url, options)
      });

      const env = buildRuntimeEnv({
        baseEnv: process.env,
        home: place,
        gatewayUrl: gateway.url,
        gatewayKey: gateway.key,
        model: job.model,
        system: job.system,
        contextPath,
        dshBin: runtimePaths.dshBin
      });
      const node = nodeExecutable();
      Object.assign(env, node.extraEnv);
      emit({ type: "step", label: "正在调用模型" });
      const text = await runDsh({
        ...node,
        dshBin: runtimePaths.dshBin,
        patchPath: runtimePaths.patchPath,
        contextPatchPath,
        cwd: place,
        env,
        model: job.model,
        prompt: job.prompt,
        onEvent: (event) => {
          if (event?.type === "step" && typeof event.label === "string") emit(event);
        }
      });
      emit({ type: "step", label: "处理完成" });
      return {
        text,
        usage: gateway.usage,
        model: job.model,
        max_tokens_clamped: gateway.max_tokens_clamped,
        finish_reasons: gateway.finish_reasons,
        truncated: gateway.truncated
      };
    } catch (error) {
      emit({ type: "step", label: error.message || "桌面执行失败" });
      throw error;
    } finally {
      if (gateway) await gateway.close().catch(() => {});
      await fs.rm(place, { recursive: true, force: true });
    }
  }

  return {
    clearAllCredentials,
    clearCredentialsForOrigin,
    models,
    prefetchAfterLogin,
    runAgent
  };
}

module.exports = { createDesktopAgent, nodeAtLeast, validateBaseUrl };
