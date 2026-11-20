const { net, safeStorage } = require("electron");
const { createHash, randomUUID } = require("node:crypto");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { createLoopbackGateway } = require("./lib/loopback-gateway.cjs");
const { runDsh } = require("./lib/dsh-bridge.cjs");
const { DEADLINE_MS, STAGE_TIMEOUT_MS, runPipeline } = require("./lib/pipeline.cjs");
const { listModelsWithRefresh, modelServiceCode, modelServiceError, parseModels } = require("./lib/models.cjs");
const { buildRuntimeEnv, resolveRuntimePaths } = require("./lib/runtime-env.cjs");
const { sessionFetch } = require("./lib/session-cookie.cjs");
const { quotaExhaustionMessage, requestUsage, USAGE_TIMEOUT_MS } = require("./lib/usage.cjs");
const { createImageGenerator } = require("./lib/image-gen.cjs");
const { createImageBroker } = require("./lib/image-broker.cjs");

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

function responseError(payload, status) {
  if (typeof payload?.detail === "string" && payload.detail.trim()) return payload.detail.trim();
  if (status === 401) return "登录已失效，请重新登录";
  return `工作台请求失败(HTTP ${status})`;
}

const PIPELINE_ROLES = ["understand", "execute", "accept", "review", "light"];
const REQUIRED_PIPELINE_ROLES = ["understand", "execute", "accept"];
const PIPELINE_ROLE_LABELS = { understand: "理解", execute: "执行", accept: "验收" };
const IMAGE_TOOL_INSTRUCTIONS = "生图工具使用说明：当制作 PPT 需要封面主视觉或关键配图，或用户明确需要图片时调用 office_image_generate（按张计费，遵守本轮张数上限）；制作 PPT 时提示词强调画面无文字，并用返回的 image_index 填入 deck.slides 的 bg_image 或 image 元素；不要在 reply 文字中输出图片链接、Markdown 图片或 base64；纯文字问答不要调用。";

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function validPipelineModel(value) {
  return typeof value === "string" && value.length >= 1 && value.length <= 200 && !/[\s\x00-\x1f\x7f]/u.test(value);
}

function sanitizePipelineConfig(payload) {
  if (!isObject(payload) || typeof payload.enabled !== "boolean" ||
      typeof payload.version !== "string" || payload.version.length < 1 || payload.version.length > 32 ||
      !isObject(payload.roles) || !isObject(payload.limits) || !isObject(payload.prompts) ||
      payload.limits.max_exec_rounds !== 2 || payload.limits.deadline_ms !== 270_000) {
    throw new Error("智能分工配置无效");
  }
  const roles = {};
  for (const role of PIPELINE_ROLES) {
    if (!validPipelineModel(payload.roles[role])) throw new Error("智能分工配置无效");
    roles[role] = payload.roles[role];
  }
  const prompts = {};
  for (const name of ["understand", "accept", "review", "revise"]) {
    const value = payload.prompts[name];
    if (typeof value !== "string" || !value.length || value.length > 8_000) {
      throw new Error("智能分工配置无效");
    }
    prompts[name] = value;
  }
  return {
    enabled: payload.enabled,
    version: payload.version,
    roles,
    limits: { max_exec_rounds: 2, deadline_ms: 270_000 },
    prompts
  };
}

function unavailableRole(config, models, role) {
  if (models.includes(config.roles[role])) return null;
  return `智能分工不可用：未开放模型 ${config.roles[role]}（角色：${PIPELINE_ROLE_LABELS[role]}）`;
}

function createDesktopAgent({ app, getWindow, getOrigin, emit = () => {}, dependencies = {} }) {
  const credentialRoot = path.join(app.getPath("userData"), "qo-model-credentials");
  const runtimePaths = resolveRuntimePaths(app.isPackaged ? process.resourcesPath : undefined, __dirname);
  const activeRuns = new Set();

  function cancelAgent() {
    let cancelled = false;
    for (const controller of activeRuns) {
      if (!controller.signal.aborted) {
        controller.abort();
        cancelled = true;
      }
    }
    return cancelled;
  }

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
    if (!response.ok) throw new Error(responseError(payload, response.status));
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
    let response;
    try {
      response = await net.fetch(url, {
        headers: { Authorization: `Bearer ${credential.apiKey}`, Accept: "application/json" },
        redirect: "error",
        signal: AbortSignal.timeout(USAGE_TIMEOUT_MS)
      });
    } catch (error) {
      if (error?.name === "TimeoutError" || error?.code === "ETIMEDOUT" || error?.code === "UND_ERR_CONNECT_TIMEOUT") {
        throw new Error("模型服务响应超时，请稍后重试");
      }
      throw new Error("无法连接模型服务，请检查网络连接后重试");
    }
    if (!response.ok) {
      let payload = {};
      try { payload = await response.json(); } catch {}
      throw Object.assign(new Error(modelServiceError(response.status, payload)), {
        status: response.status,
        code: modelServiceCode(payload)
      });
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

  const readFreshModels = dependencies.freshModels || freshModels;
  const getCredentialImpl = dependencies.getCredential || getCredential;
  const readSessionJson = dependencies.fetchSessionJson || fetchSessionJson;
  const makeGateway = dependencies.createLoopbackGateway || createLoopbackGateway;
  const runDshStage = dependencies.runDsh || runDsh;
  const readUsage = dependencies.requestUsage || requestUsage;
  const clockNow = dependencies.now || Date.now;
  const getNodeExecutable = dependencies.nodeExecutable || nodeExecutable;
  const generateImage = createImageGenerator({
    getModels: readFreshModels,
    getCredential: async () => (await getCredentialImpl(getOrigin())).credential,
    fetchImpl: dependencies.imageFetch || ((url, options) => net.fetch(url, options)),
    now: clockNow,
    timeoutMs: dependencies.imageTimeoutMs || 180_000
  });
  const imageGenerateImpl = dependencies.imageGenerate || generateImage;
  const createImageBrokerImpl = dependencies.createImageBroker || createImageBroker;

  async function loadPipelineConfig(origin) {
    const payload = await readSessionJson(origin, `${origin}/api/v1/me/pipeline`);
    return sanitizePipelineConfig(payload);
  }

  async function pipelineStatus(origin, available) {
    try {
      const config = await loadPipelineConfig(origin);
      if (!config.enabled) return { available: false, reason: "智能分工未开放", roles: config.roles };
      const missing = REQUIRED_PIPELINE_ROLES.map((role) => unavailableRole(config, available, role)).find(Boolean);
      if (missing) return { available: false, reason: missing, roles: config.roles };
      return { available: true, roles: config.roles };
    } catch {
      return { available: false, reason: "智能分工配置获取失败" };
    }
  }

  async function models() {
    try {
      const available = await readFreshModels();
      const pipeline = await pipelineStatus(getOrigin(), available);
      const image_models = available.filter((model) => model.startsWith("gpt-image-")).sort((a, b) => a.localeCompare(b));
      return { models: available.filter((model) => !model.startsWith("gpt-image-")), image_models, pipeline };
    } catch (error) {
      return { models: [], image_models: [], error: error.message || "模型列表加载失败",
        pipeline: { available: false, reason: "模型列表获取失败" } };
    }
  }

  async function usage() {
    try {
      const { credential } = await getCredential();
      return await requestUsage(credential, (url, options) => net.fetch(url, options));
    } catch {
      return { state: "unavailable" };
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
    const controller = new AbortController();
    const { signal } = controller;
    activeRuns.add(controller);
    const ensureNotAborted = () => {
      if (signal.aborted) throw Object.assign(new Error("已停止生成"), { code: "ABORTED" });
    };
    const startedAt = clockNow();
    const pipelineRequested = job.model === "qo:pipeline";
    const origin = getOrigin();
    const hasContext = job.context !== undefined && job.context !== null;
    try {
    emit({ type: "step", label: hasContext ? "载入项目资料" : "准备执行环境" });
    const available = await readFreshModels();
    ensureNotAborted();
    let config;
    if (pipelineRequested) {
      try {
        config = await loadPipelineConfig(origin);
      } catch {
        throw new Error("智能分工配置获取失败");
      }
      if (!config.enabled) throw new Error("智能分工未开放");
      for (const role of REQUIRED_PIPELINE_ROLES) {
        const missing = unavailableRole(config, available, role);
        if (missing) throw new Error(missing);
      }
    } else if (!available.includes(job.model)) {
      throw new Error("该模型已不再对你开放，请重新选择");
    } else if (job.model.startsWith("gpt-image-")) {
      throw new Error("生图模型只能用于生图");
    }
    const imageModels = available.filter((model) => model.startsWith("gpt-image-")).sort((a, b) => a.localeCompare(b));
    const imageToolEnabled = job.image_tool === true && imageModels.length > 0;
    const maxImages = Number.isInteger(job.max_images) && job.max_images >= 1 && job.max_images <= 3 ? job.max_images : 2;
    const imageModel = imageModels.includes("gpt-image-2.6-flare") ? "gpt-image-2.6-flare" : imageModels[0];
    ensureNotAborted();
    const { credential } = await getCredentialImpl(origin);
    ensureNotAborted();
    let imageBroker;

    async function runStage(stage) {
      ensureNotAborted();
      const place = await fs.mkdtemp(path.join(os.tmpdir(), "qo-office-agent-"));
      let gateway;
      const stageImageEnabled = Boolean(imageToolEnabled && imageBroker && stage.withImages !== false);
      try {
        let contextPath;
        let contextPatchPath;
        const extraPatchPaths = [];
        if (stage.withContext && hasContext) {
          contextPath = path.join(place, "office-context.json");
          await fs.writeFile(contextPath, JSON.stringify(job.context), { encoding: "utf8", mode: 0o600 });
          contextPatchPath = path.join(place, "office-context.patch.yml");
          await fs.writeFile(contextPatchPath,
            `- insert:\n    - id: office-context\n      name: ${JSON.stringify(runtimePaths.contextPlugin)}\n`,
            { encoding: "utf8", mode: 0o600 });
        }
        if (stageImageEnabled) {
          const imagePatchPath = path.join(place, "office-image.patch.yml");
          const pluginPath = path.join(runtimePaths.harnessRoot, "office-image.mjs").replaceAll("\\", "/");
          await fs.writeFile(imagePatchPath,
            `- insert:\n    - id: office-image\n      name: ${JSON.stringify(pluginPath)}\n`,
            { encoding: "utf8", mode: 0o600 });
          extraPatchPaths.push(imagePatchPath);
        }

        gateway = await makeGateway({
          baseUrl: credential.baseUrl,
          apiKey: credential.apiKey,
          model: stage.model,
          allowedTools: [
            ...(contextPath ? ["office_context_read"] : []),
            ...(stageImageEnabled ? ["office_image_generate"] : [])
          ],
          stickyQuotaError: pipelineRequested,
          fetchImpl: (url, options) => net.fetch(url, options)
        });

        const env = buildRuntimeEnv({
          baseEnv: process.env,
          home: place,
          gatewayUrl: gateway.url,
          gatewayKey: gateway.key,
          model: stage.model,
          system: stage.system,
          contextPath,
          ...(stageImageEnabled ? {
            imageBrokerUrl: imageBroker.url,
            imageBrokerKey: imageBroker.key
          } : {}),
          dshBin: runtimePaths.dshBin
        });
        const node = getNodeExecutable();
        Object.assign(env, node.extraEnv);
        emit({ type: "step", label: "正在调用模型" });
        const text = await runDshStage({
          ...node,
          dshBin: runtimePaths.dshBin,
          patchPath: runtimePaths.patchPath,
          contextPatchPath,
          ...(extraPatchPaths.length ? { extraPatchPaths } : {}),
          cwd: place,
          env,
          model: stage.model,
          prompt: stage.prompt,
          timeoutMs: stage.timeoutMs,
          signal,
          onEvent: (event) => {
            if (event?.type === "step" && typeof event.label === "string") emit(event);
          }
        });
        if (pipelineRequested && gateway?.quota_error) {
          const failure = new Error("模型额度已用尽");
          failure.quota_error = gateway.quota_error;
          throw failure;
        }
        return {
          text,
          usage: {
            ...gateway.usage,
            max_tokens_clamped: gateway.max_tokens_clamped,
            finish_reasons: gateway.finish_reasons,
            truncated: gateway.truncated
          }
        };
      } catch (error) {
        if (error?.code === "ABORTED" || signal.aborted) {
          throw Object.assign(new Error("已停止生成"), { code: "ABORTED" });
        }
        if (gateway?.quota_error) {
          let usageSnapshot;
          try {
            const result = await readUsage(credential, (url, options) => net.fetch(url, options));
            if (result.state === "available") usageSnapshot = result.usage;
          } catch {}
          const message = quotaExhaustionMessage(gateway.quota_error, usageSnapshot);
          const failure = new Error(message || error.message || "模型额度已用尽");
          failure.quota_error = gateway.quota_error;
          throw failure;
        }
        if (/桌面 Harness 调用超时/.test(error?.message || "")) error.code = "STAGE_TIMEOUT";
        throw error;
      } finally {
        if (gateway) await gateway.close().catch(() => {});
        await fs.rm(place, { recursive: true, force: true });
      }
    }

    try {
      if (imageToolEnabled) {
        imageBroker = await createImageBrokerImpl({
          generate: (request, options) => imageGenerateImpl(request, options),
          model: imageModel,
          maxImages,
          deadlineAt: startedAt + (pipelineRequested ? DEADLINE_MS : 600_000),
          now: clockNow
        });
      }
      if (pipelineRequested) {
        const pipelineJob = imageBroker
          ? { ...job, system: `${job.system}\n\n${IMAGE_TOOL_INSTRUCTIONS}` }
          : job;
        const pipelineResult = await runPipeline({
          job: pipelineJob,
          config,
          availableModels: available,
          runStage,
          emit,
          now: clockNow,
          startedAt,
          signal
        });
        return {
          ...pipelineResult,
          ...(imageBroker?.images.length ? { images: imageBroker.images.map((image) => ({ ...image })) } : {})
        };
      }
      const result = await runStage({
        role: "execute",
        model: job.model,
        system: imageBroker ? `${job.system}\n\n${IMAGE_TOOL_INSTRUCTIONS}` : job.system,
        prompt: job.prompt,
        withContext: hasContext,
        timeoutMs: STAGE_TIMEOUT_MS
      });
      emit({ type: "step", label: "处理完成" });
      return {
        text: result.text,
        usage: result.usage,
        model: job.model,
        max_tokens_clamped: result.usage.max_tokens_clamped,
        finish_reasons: result.usage.finish_reasons,
        truncated: result.usage.truncated,
        ...(imageBroker ? { images: imageBroker.images.map((image) => ({ ...image })) } : {})
      };
    } catch (error) {
      if (imageBroker?.images.length) {
        error.message = `${error.message || "桌面执行失败"}（本轮已生成 ${imageBroker.images.length} 张图片并已计费，因执行失败未能保存）`;
      }
      emit({ type: "step", label: error.message || "桌面执行失败" });
      throw error;
    } finally {
      await imageBroker?.close().catch(() => {});
    }
    } finally {
      activeRuns.delete(controller);
    }
  }

  return {
    cancelAgent,
    clearAllCredentials,
    clearCredentialsForOrigin,
    models,
    usage,
    generateImage,
    prefetchAfterLogin,
    runAgent
  };
}

module.exports = { createDesktopAgent, nodeAtLeast, responseError, sanitizePipelineConfig, validateBaseUrl };
