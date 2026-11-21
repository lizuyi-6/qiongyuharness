const { net, safeStorage } = require("electron");
const { createHash, randomUUID } = require("node:crypto");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { createLoopbackGateway } = require("./lib/loopback-gateway.cjs");
const { runDsh } = require("./lib/dsh-bridge.cjs");
const { DEADLINE_MS, STAGE_TIMEOUT_MS, buildResumeContext, runPipeline } = require("./lib/pipeline.cjs");
const { createCheckpointStore, describeStage } = require("./lib/checkpoint.cjs");
const { listModelsWithRefresh, modelServiceCode, modelServiceError, parseModels } = require("./lib/models.cjs");
const { buildRuntimeEnv, resolveRuntimePaths } = require("./lib/runtime-env.cjs");
const { sessionFetch } = require("./lib/session-cookie.cjs");
const { quotaExhaustionMessage, requestUsage, USAGE_TIMEOUT_MS } = require("./lib/usage.cjs");
const { createImageGenerator } = require("./lib/image-gen.cjs");
const { createImageBroker } = require("./lib/image-broker.cjs");
const { createSearchBroker, normalizeServerTime } = require("./lib/search-broker.cjs");

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
const IMAGE_TOOL_INSTRUCTIONS = "生图工具使用说明：当制作 PPT 需要封面主视觉、章节过渡或关键配图，或用户明确需要图片时，可按需调用 office_image_generate（按张计费，不限张数）；制作 PPT 时提示词强调画面无文字，并用返回的 image_index 填入 deck.slides 的 bg_image 或 image 元素；不要在 reply 文字中输出图片链接、Markdown 图片或 base64；不需要配图的纯文字问答不要调用。";
const SEARCH_TOOL_INSTRUCTIONS = "DeepSeek 联网搜索工具使用说明（强制要求）：本轮已提供 web_search 与 web_fetch 工具。在给出最终答复、文档或 PPT 之前，你必须至少调用 1 次 web_search（传入 1–4 条关键词）检索最新权威资料，需要精读某条来源时调用 web_fetch，并在回答中注明来源链接。未调用 web_search 将被程序拦截重写。";

function buildSearchToolInstructions(serverTime) {
  if (serverTime && serverTime.year && serverTime.date && serverTime.unix_timestamp) {
    return `DeepSeek 联网搜索工具使用说明（强制按当前服务器时间戳检索）：当前服务器权威时间为 ${serverTime.display || serverTime.date}（Unix时间戳：${serverTime.unix_timestamp}，当前真实年份：${serverTime.year}年）。本轮已提供 web_search 与 web_fetch 工具。在给出最终答复、文档或 PPT 之前，你必须先根据当前服务器时间戳（${serverTime.year}年）至少调用 1 次 web_search（传入 1–4 条关键词，检索最新动态/产品/模型/数据时必须显式包含 ${serverTime.year}，严禁把 2024 或 2025 当作当前年份！），需要精读某条来源时调用 web_fetch，并在回答中采用离当前时间戳最近的 ${serverTime.year} 年最新权威事实。未调用 web_search 将被程序拦截重写。`;
  }
  return SEARCH_TOOL_INSTRUCTIONS;
}

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
  const createSearchBrokerImpl = dependencies.createSearchBroker || createSearchBroker;
  const checkpointStore = dependencies.checkpointStore || createCheckpointStore({
    baseDir: path.join(app.getPath("userData"), "qo-checkpoints"),
    now: clockNow
  });

  async function getCheckpoint(conversationId, userId = "") {
    const origin = getOrigin();
    if (!origin || !conversationId) return null;
    return checkpointStore.summary(origin, String(conversationId), { userId: String(userId || "") });
  }

  async function clearCheckpoint(conversationId) {
    const origin = getOrigin();
    if (!origin || !conversationId) return false;
    return checkpointStore.clear(origin, String(conversationId));
  }

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
    const conversationId = typeof job.conversation_id === "string" ? job.conversation_id : "";
    const resumeRequested = job.resume === true;
    const hasContext = job.context !== undefined && job.context !== null;
    const recordedSteps = [];
    const emitStep = (event) => {
      if (event?.type === "step" && typeof event.label === "string" && event.label) {
        if (recordedSteps[recordedSteps.length - 1] !== event.label) {
          recordedSteps.push(event.label);
          if (recordedSteps.length > 60) recordedSteps.shift();
        }
      }
      emit(event);
    };

    let resumedCheckpoint = null;
    if (resumeRequested && origin && conversationId) {
      resumedCheckpoint = await checkpointStore.load(origin, conversationId, { userId: job.user_id || "" });
    }
    if (resumedCheckpoint?.stage === "complete" &&
        resumedCheckpoint.completedResult &&
        typeof resumedCheckpoint.completedResult.text === "string" &&
        resumedCheckpoint.completedResult.text.trim()) {
      emitStep({ type: "step", label: "断点续跑：复用已生成的完整结果，正在重新提交保存" });
      return {
        ...resumedCheckpoint.completedResult,
        resumedFromComplete: true
      };
    }

    const checkpointMeta = job.checkpoint_meta && typeof job.checkpoint_meta === "object" ? job.checkpoint_meta : {};
    const checkpointState = {
      runId: typeof job.run_id === "string" ? job.run_id : (resumedCheckpoint?.runId || ""),
      userId: typeof job.user_id === "string" ? job.user_id : (resumedCheckpoint?.userId || ""),
      message: typeof checkpointMeta.message === "string" ? checkpointMeta.message : (resumedCheckpoint?.message || ""),
      mode: typeof checkpointMeta.mode === "string" ? checkpointMeta.mode : (resumedCheckpoint?.mode || "auto"),
      materialIds: Array.isArray(checkpointMeta.material_ids) ? checkpointMeta.material_ids : (resumedCheckpoint?.materialIds || []),
      selection: checkpointMeta.selection !== undefined ? checkpointMeta.selection : (resumedCheckpoint?.selection || null),
      model: job.model,
      allowImageTool: job.image_tool === true,
      stage: resumedCheckpoint?.stage || (pipelineRequested ? "understand" : "single_execute"),
      brief: resumedCheckpoint?.brief || null,
      tier: resumedCheckpoint?.tier || "standard",
      finalText: resumedCheckpoint?.finalText || "",
      checkedDeck: resumedCheckpoint?.checkedDeck || null,
      verdict: resumedCheckpoint?.verdict || "fail",
      review: resumedCheckpoint?.review || "not_needed",
      rounds: resumedCheckpoint?.rounds || 1,
      issues: Array.isArray(resumedCheckpoint?.issues) ? [...resumedCheckpoint.issues] : [],
      stages: Array.isArray(resumedCheckpoint?.stages) ? [...resumedCheckpoint.stages] : [],
      twoRoundAcceptanceFailed: Boolean(resumedCheckpoint?.twoRoundAcceptanceFailed),
      searchState: resumedCheckpoint?.searchState || null,
      images: Array.isArray(resumedCheckpoint?.images) ? resumedCheckpoint.images.map((img) => ({ ...img })) : [],
      thinkingTail: typeof resumedCheckpoint?.thinkingTail === "string" ? resumedCheckpoint.thinkingTail : "",
      steps: Array.isArray(resumedCheckpoint?.steps) ? [...resumedCheckpoint.steps] : [],
      error: "",
      completedResult: null
    };
    const initialUploadedImageCount = checkpointState.images.length;
    let latestThinkingTail = checkpointState.thinkingTail;

    const persistCheckpoint = (patch = {}) => {
      if (!origin || !conversationId) return Promise.resolve(null);
      Object.assign(checkpointState, patch);
      if (recordedSteps.length) {
        const merged = [...checkpointState.steps];
        for (const s of recordedSteps) {
          if (merged[merged.length - 1] !== s) merged.push(s);
        }
        checkpointState.steps = merged.slice(-60);
      }
      checkpointState.thinkingTail = latestThinkingTail;
      return checkpointStore.save(origin, conversationId, checkpointState).catch(() => null);
    };

    try {
    emitStep({ type: "step", label: resumedCheckpoint ? `从「${describeStage(resumedCheckpoint.stage)}」断点续跑` : (hasContext ? "载入项目资料" : "准备执行环境") });
    if (origin && conversationId) void persistCheckpoint();
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
    const pptCheckEnabled = job.ppt_check === true;
    let webSearchEnabled = job.web_search?.enabled === true;
    let webSearchRequired = job.web_search?.required === true;
    let rawServerTime = job.web_search?.server_time || job.server_time || null;
    if ((job.web_search === undefined || !rawServerTime) && origin) {
      try {
        const searchStatus = await readSessionJson(origin, `${origin}/api/v1/search/status`);
        if (job.web_search === undefined) {
          webSearchEnabled = searchStatus?.configured === true;
          webSearchRequired = Boolean(webSearchEnabled && searchStatus?.required !== false);
        }
        if (!rawServerTime && searchStatus?.server_time) {
          rawServerTime = searchStatus.server_time;
        }
      } catch {}
    }
    const serverTime = normalizeServerTime(rawServerTime, clockNow);
    const maxImages = Number.isInteger(job.max_images) && job.max_images >= 1 ? job.max_images : null;
    const imageModel = imageModels.includes("gpt-image-2.6-flare") ? "gpt-image-2.6-flare" : imageModels[0];
    ensureNotAborted();
    const { credential } = await getCredentialImpl(origin);
    ensureNotAborted();
    let imageBroker;
    let searchBroker;

    const toolDisplayNames = {
      office_context_read: "读取授权项目资料",
      office_image_generate: "生成专属配图",
      office_ppt_check: "校验 PPT 自由画布排版",
      web_search: "联网搜索（DeepSeek Search）",
      web_fetch: "读取网页全文"
    };

    async function runStage(stage) {
      ensureNotAborted();
      const place = await fs.mkdtemp(path.join(os.tmpdir(), "qo-office-agent-"));
      let gateway;
      const stageImageEnabled = Boolean(imageToolEnabled && imageBroker && stage.withImages !== false);
      const stagePptCheckEnabled = Boolean(pptCheckEnabled && stage.withImages !== false);
      const stageSearchEnabled = Boolean(webSearchEnabled && searchBroker && stage.withSearch !== false);
      let thinkingBuffer = "";
      let thinkingTimer = null;
      let contentChars = 0;
      let lastProgressAt = 0;
      const flushThinking = () => {
        if (thinkingTimer) {
          clearTimeout(thinkingTimer);
          thinkingTimer = null;
        }
        if (!thinkingBuffer) return;
        const text = thinkingBuffer;
        thinkingBuffer = "";
        latestThinkingTail = `${latestThinkingTail}${text}`.slice(-3200);
        emit({ type: "thinking", text, stage: stage.role, model: stage.model });
      };
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
        if (stagePptCheckEnabled) {
          const checkPatchPath = path.join(place, "office-ppt-check.patch.yml");
          const checkPluginPath = path.join(runtimePaths.harnessRoot, "office-ppt-check.mjs").replaceAll("\\", "/");
          await fs.writeFile(checkPatchPath,
            `- insert:\n    - id: office-ppt-check\n      name: ${JSON.stringify(checkPluginPath)}\n`,
            { encoding: "utf8", mode: 0o600 });
          extraPatchPaths.push(checkPatchPath);
        }
        if (stageSearchEnabled) {
          const searchPatchPath = path.join(place, "office-web-search.patch.yml");
          const searchPluginPath = path.join(runtimePaths.harnessRoot, "office-web-search.mjs").replaceAll("\\", "/");
          await fs.writeFile(searchPatchPath,
            `- insert:\n    - id: office-web-search\n      name: ${JSON.stringify(searchPluginPath)}\n`,
            { encoding: "utf8", mode: 0o600 });
          extraPatchPaths.push(searchPatchPath);
        }

        gateway = await makeGateway({
          baseUrl: credential.baseUrl,
          apiKey: credential.apiKey,
          model: stage.model,
          allowedTools: [
            ...(contextPath ? ["office_context_read"] : []),
            ...(stageImageEnabled ? ["office_image_generate"] : []),
            ...(stagePptCheckEnabled ? ["office_ppt_check"] : []),
            ...(stageSearchEnabled ? ["web_search", "web_fetch"] : [])
          ],
          stickyQuotaError: pipelineRequested,
          fetchImpl: (url, options) => net.fetch(url, options),
          onStreamEvent: (streamEvent) => {
            if (!streamEvent) return;
            if (streamEvent.type === "thinking" && typeof streamEvent.text === "string" && streamEvent.text) {
              thinkingBuffer += streamEvent.text;
              if (thinkingBuffer.length >= 160) {
                flushThinking();
              } else if (!thinkingTimer) {
                thinkingTimer = setTimeout(flushThinking, 120);
                thinkingTimer.unref?.();
              }
            } else if (streamEvent.type === "tool" && typeof streamEvent.name === "string") {
              flushThinking();
              const label = toolDisplayNames[streamEvent.name] || streamEvent.name;
              emitStep({ type: "step", label: `调用工具：${label}` });
            } else if (streamEvent.type === "content" && Number.isFinite(streamEvent.chars)) {
              flushThinking();
              contentChars += streamEvent.chars;
              const nowMs = Date.now();
              if (contentChars > 0 && (lastProgressAt === 0 || nowMs - lastProgressAt >= 550)) {
                lastProgressAt = nowMs;
                emit({ type: "progress", chars: contentChars, stage: stage.role, model: stage.model });
              }
            }
          }
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
          ...(stageSearchEnabled ? {
            searchBrokerUrl: searchBroker.url,
            searchBrokerKey: searchBroker.key
          } : {}),
          serverTime: searchBroker?.serverTime || serverTime,
          dshBin: runtimePaths.dshBin
        });
        const node = getNodeExecutable();
        Object.assign(env, node.extraEnv);
        emitStep({ type: "step", label: "正在调用模型" });
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
            if (event?.type === "step" && typeof event.label === "string") emitStep(event);
          },
          onCheckedDeck: (deck) => {
            if (deck && typeof deck === "object" && Array.isArray(deck.slides) && deck.slides.length) {
              void persistCheckpoint({ checkedDeck: deck });
            }
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
        flushThinking();
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
          deadlineAt: startedAt + (pipelineRequested ? DEADLINE_MS : STAGE_TIMEOUT_MS),
          now: clockNow,
          initialImages: checkpointState.images,
          onImageGenerated: (_img, _index, allImages) => {
            void persistCheckpoint({ images: allImages });
          }
        });
      }
      if (webSearchEnabled) {
        searchBroker = await createSearchBrokerImpl({
          search: (payload, { signal: reqSignal }) => readSessionJson(origin, `${origin}/api/v1/search/query`, {
            method: "POST",
            headers: { "Content-Type": "application/json", Accept: "application/json" },
            body: JSON.stringify(payload),
            signal: reqSignal
          }),
          fetchPage: (payload, { signal: reqSignal }) => readSessionJson(origin, `${origin}/api/v1/search/fetch`, {
            method: "POST",
            headers: { "Content-Type": "application/json", Accept: "application/json" },
            body: JSON.stringify(payload),
            signal: reqSignal
          }),
          serverTime,
          now: clockNow,
          initialState: checkpointState.searchState,
          onSearch: (queries, srvTime) => {
            const label = queries.slice(0, 3).join(" / ").slice(0, 100);
            const dateTag = srvTime?.date || serverTime?.date || "";
            emitStep({ type: "step", label: `DeepSeek 搜索${dateTag ? `（基准 ${dateTag}）` : ""}：${label}` });
          },
          onFetch: (targetUrl) => {
            let host = targetUrl;
            try { host = new URL(targetUrl).hostname; } catch {}
            emitStep({ type: "step", label: `读取网页：${host}` });
          },
          onStateChange: (snapshot) => {
            void persistCheckpoint({ searchState: snapshot });
          }
        });
      }
      const extraSystemParts = [];
      if (imageBroker) extraSystemParts.push(IMAGE_TOOL_INSTRUCTIONS);
      if (searchBroker && webSearchRequired) extraSystemParts.push(buildSearchToolInstructions(searchBroker.serverTime || serverTime));
      const effectiveSystem = extraSystemParts.length ? `${job.system}\n\n${extraSystemParts.join("\n\n")}` : job.system;
      const collectImages = () => {
        const list = imageBroker ? imageBroker.images : checkpointState.images;
        return Array.isArray(list) ? list.map((image) => ({ ...image })) : [];
      };
      if (pipelineRequested) {
        const pipelineJob = { ...job, system: effectiveSystem };
        const pipelineResumeState = resumedCheckpoint
          ? {
              stage: resumedCheckpoint.stage,
              stageLabel: describeStage(resumedCheckpoint.stage),
              error: resumedCheckpoint.error || "",
              brief: resumedCheckpoint.brief || null,
              tier: resumedCheckpoint.tier || "standard",
              finalText: resumedCheckpoint.finalText || "",
              checkedDeck: resumedCheckpoint.checkedDeck || null,
              verdict: resumedCheckpoint.verdict || "fail",
              review: resumedCheckpoint.review || "not_needed",
              rounds: resumedCheckpoint.rounds || 1,
              issues: Array.isArray(resumedCheckpoint.issues) ? resumedCheckpoint.issues : [],
              stages: Array.isArray(resumedCheckpoint.stages) ? resumedCheckpoint.stages : [],
              twoRoundAcceptanceFailed: Boolean(resumedCheckpoint.twoRoundAcceptanceFailed),
              thinkingTail: resumedCheckpoint.thinkingTail || ""
            }
          : null;
        const pipelineResult = await runPipeline({
          job: pipelineJob,
          config,
          availableModels: available,
          runStage,
          emit: emitStep,
          now: clockNow,
          startedAt,
          signal,
          getGeneratedImageCount: () => collectImages().length,
          getGeneratedImages: collectImages,
          requireSearch: Boolean(searchBroker && webSearchRequired),
          getSearchCount: () => searchBroker?.searchCount || 0,
          serverTime: searchBroker?.serverTime || serverTime,
          getSearchEvidence: () => searchBroker?.getEvidenceDigest?.() || "",
          resumeState: pipelineResumeState,
          onCheckpoint: (snapshot) => {
            void persistCheckpoint({
              ...snapshot,
              images: collectImages(),
              ...(searchBroker?.exportState ? { searchState: searchBroker.exportState() } : {})
            });
          }
        });
        const finalImages = collectImages();
        const finalPayload = {
          ...pipelineResult,
          ...(finalImages.length ? { images: finalImages } : {})
        };
        await persistCheckpoint({
          stage: "complete",
          finalText: finalPayload.text,
          images: finalImages,
          ...(searchBroker?.exportState ? { searchState: searchBroker.exportState() } : {}),
          completedResult: finalPayload,
          error: ""
        });
        return finalPayload;
      }
      const singleResumeExtra = resumedCheckpoint
        ? buildResumeContext({
            resumeState: {
              stage: resumedCheckpoint.stage,
              stageLabel: describeStage(resumedCheckpoint.stage),
              error: resumedCheckpoint.error || "",
              checkedDeck: resumedCheckpoint.checkedDeck || null,
              thinkingTail: resumedCheckpoint.thinkingTail || ""
            },
            searchEvidence: searchBroker?.getEvidenceDigest?.() || "",
            images: collectImages()
          })
        : "";
      let result = await runStage({
        role: "execute",
        model: job.model,
        system: `${effectiveSystem}${singleResumeExtra}`,
        prompt: job.prompt,
        withContext: hasContext,
        withSearch: Boolean(searchBroker),
        timeoutMs: STAGE_TIMEOUT_MS
      });
      if (searchBroker && webSearchRequired && searchBroker.searchCount < 1) {
        emitStep({ type: "step", label: "程序校验：尚未联网搜索，正在强制调用 DeepSeek 搜索" });
        result = await runStage({
          role: "execute",
          model: job.model,
          system: `${effectiveSystem}${singleResumeExtra}`,
          prompt: `${job.prompt}\n\n【程序强制拦截】本轮强制要求先调用 web_search（DeepSeek 联网搜索）检索最新资料后再作答。你刚才未调用 web_search，请立即调用 web_search（传入 1–4 个关键词），基于检索结果重新输出完整 JSON 交付物！`,
          withContext: hasContext,
          withImages: !(collectImages().length),
          withSearch: true,
          timeoutMs: STAGE_TIMEOUT_MS
        });
      }
      emitStep({ type: "step", label: "处理完成" });
      const finalImages = collectImages();
      const finalPayload = {
        text: result.text,
        usage: result.usage,
        model: job.model,
        max_tokens_clamped: result.usage.max_tokens_clamped,
        finish_reasons: result.usage.finish_reasons,
        truncated: result.usage.truncated,
        ...(imageBroker || finalImages.length ? { images: finalImages } : {})
      };
      await persistCheckpoint({
        stage: "complete",
        finalText: finalPayload.text,
        images: finalImages,
        ...(searchBroker?.exportState ? { searchState: searchBroker.exportState() } : {}),
        completedResult: finalPayload,
        error: ""
      });
      return finalPayload;
    } catch (error) {
      const currentImages = imageBroker ? imageBroker.images : checkpointState.images;
      const newImagesToSave = imageBroker ? imageBroker.images.slice(initialUploadedImageCount) : [];
      if (error?.checkedDeck && typeof error.checkedDeck === "object") {
        checkpointState.checkedDeck = error.checkedDeck;
      }
      await persistCheckpoint({
        error: String(error?.message || "桌面执行失败").slice(0, 300),
        images: Array.isArray(currentImages) ? currentImages.map((image) => ({ ...image })) : [],
        ...(searchBroker?.exportState ? { searchState: searchBroker.exportState() } : {})
      });
      if (newImagesToSave.length && origin && job.conversation_id && job.run_id) {
        const imgEndpoint = `${origin}/api/v1/conversations/${encodeURIComponent(job.conversation_id)}/turns/${encodeURIComponent(job.run_id)}/images`;
        for (const image of newImagesToSave) {
          try {
            await readSessionJson(origin, imgEndpoint, {
              method: "POST",
              headers: { "Content-Type": "application/json", Accept: "application/json" },
              body: JSON.stringify(image)
            });
          } catch {}
        }
      }
      emitStep({ type: "step", label: error.message || "桌面执行失败" });
      throw error;
    } finally {
      await imageBroker?.close().catch(() => {});
      await searchBroker?.close().catch(() => {});
    }
    } finally {
      activeRuns.delete(controller);
    }
  }

  return {
    cancelAgent,
    clearAllCredentials,
    clearCheckpoint,
    clearCredentialsForOrigin,
    getCheckpoint,
    models,
    usage,
    generateImage,
    prefetchAfterLogin,
    runAgent
  };
}

module.exports = { createDesktopAgent, nodeAtLeast, responseError, sanitizePipelineConfig, validateBaseUrl };
