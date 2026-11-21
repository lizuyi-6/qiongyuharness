"use strict";

const { createHash, randomUUID } = require("node:crypto");
const fs = require("node:fs/promises");
const path = require("node:path");

const CHECKPOINT_VERSION = 1;
const MAX_CHECKPOINT_AGE_MS = 7 * 24 * 60 * 60 * 1000;
const STAGE_LABELS = {
  understand: "需求理解",
  light: "轻量执行",
  execute_round_1: "方案执行",
  accept_round_1: "独立验收",
  execute_round_2: "返工修订",
  accept_round_2: "复验确认",
  review: "高级复核",
  single_execute: "任务执行",
  complete: "结果保存"
};

function digest(value) {
  return createHash("sha256").update(String(value || "")).digest("hex");
}

function describeStage(stage) {
  return STAGE_LABELS[stage] || "任务执行";
}

function summarizeCheckpoint(data) {
  if (!data || typeof data !== "object") return null;
  const images = Array.isArray(data.images) ? data.images : [];
  const searchState = data.searchState && typeof data.searchState === "object" ? data.searchState : {};
  const evidenceCount = Array.isArray(searchState.evidenceByUrl) ? searchState.evidenceByUrl.length : 0;
  const fetchedCount = Array.isArray(searchState.fetchedPages) ? searchState.fetchedPages.length : 0;
  const searchCount = Number.isInteger(searchState.searchCount) ? searchState.searchCount : 0;
  const stagesDone = Array.isArray(data.stages) ? data.stages.filter((s) => s && s.outcome !== "format_error").length : 0;
  const hasBrief = Boolean(data.brief && typeof data.brief === "object");
  const hasDraft = Boolean(data.checkedDeck && typeof data.checkedDeck === "object" && Array.isArray(data.checkedDeck.slides) && data.checkedDeck.slides.length);
  const hasFinalText = typeof data.finalText === "string" && data.finalText.trim().length > 0;
  const hasCompletedResult = Boolean(data.stage === "complete" && data.completedResult && typeof data.completedResult.text === "string" && data.completedResult.text.trim());
  const stage = typeof data.stage === "string" && data.stage ? data.stage : "single_execute";

  const preserved = [];
  if (hasCompletedResult) {
    preserved.push("已生成完整结果（可直接保存，无需重跑模型）");
  } else {
    if (hasBrief) preserved.push("已理解需求");
    if (searchCount > 0 || evidenceCount > 0) {
      preserved.push(evidenceCount > 0 ? `已联网检索 ${evidenceCount} 条权威资料` : `已执行 ${searchCount} 次联网搜索`);
    }
    if (fetchedCount > 0) preserved.push(`已精读 ${fetchedCount} 个网页`);
    if (images.length > 0) preserved.push(`已保留 ${images.length} 张配图（续跑不重画）`);
    if (hasFinalText) preserved.push(data.rounds >= 2 ? "已完成返工稿" : "已完成初稿");
    else if (hasDraft) preserved.push(`已保留 ${data.checkedDeck.slides.length} 页 PPT 草稿`);
  }

  return {
    version: CHECKPOINT_VERSION,
    conversationId: String(data.conversationId || ""),
    runId: String(data.runId || ""),
    userId: String(data.userId || ""),
    savedAt: Number(data.savedAt) || Date.now(),
    stage,
    stageLabel: describeStage(stage),
    error: typeof data.error === "string" ? data.error : "",
    message: typeof data.message === "string" ? data.message : "",
    mode: typeof data.mode === "string" ? data.mode : "auto",
    materialIds: Array.isArray(data.materialIds) ? data.materialIds : [],
    selection: data.selection && typeof data.selection === "object" ? data.selection : null,
    model: typeof data.model === "string" ? data.model : "",
    allowImageTool: data.allowImageTool !== false,
    hasBrief,
    hasDraft,
    hasFinalText,
    hasCompletedResult,
    searchCount,
    evidenceCount,
    fetchedCount,
    imageCount: images.length,
    stagesDone,
    steps: Array.isArray(data.steps) ? data.steps.slice(-40) : [],
    preserved
  };
}

function createCheckpointStore({ baseDir, now = Date.now }) {
  if (!baseDir || typeof baseDir !== "string") {
    throw new TypeError("checkpoint baseDir is required");
  }
  const writeChains = new Map();

  function originDir(origin) {
    return path.join(baseDir, digest(origin || "default"));
  }

  function filePath(origin, conversationId) {
    return path.join(originDir(origin), `${digest(conversationId)}.json`);
  }

  function enqueue(key, task) {
    const prev = writeChains.get(key) || Promise.resolve();
    const next = prev.catch(() => {}).then(task);
    writeChains.set(key, next);
    next.finally(() => {
      if (writeChains.get(key) === next) writeChains.delete(key);
    });
    return next;
  }

  async function save(origin, conversationId, payload) {
    if (!origin || !conversationId || !payload || typeof payload !== "object") return null;
    const file = filePath(origin, conversationId);
    const record = {
      ...payload,
      version: CHECKPOINT_VERSION,
      origin,
      conversationId,
      savedAt: now()
    };
    return enqueue(file, async () => {
      await fs.mkdir(path.dirname(file), { recursive: true });
      const tmp = `${file}.${randomUUID()}.tmp`;
      await fs.writeFile(tmp, JSON.stringify(record), { encoding: "utf8", mode: 0o600 });
      await fs.rename(tmp, file);
      return record;
    });
  }

  async function load(origin, conversationId, { userId = "" } = {}) {
    if (!origin || !conversationId) return null;
    const file = filePath(origin, conversationId);
    const pending = writeChains.get(file);
    if (pending) await pending.catch(() => {});
    try {
      const raw = await fs.readFile(file, "utf8");
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object" || parsed.version !== CHECKPOINT_VERSION) return null;
      if (parsed.origin !== origin || parsed.conversationId !== conversationId) return null;
      if (userId && parsed.userId && parsed.userId !== userId) return null;
      const age = now() - (Number(parsed.savedAt) || 0);
      if (!Number.isFinite(age) || age > MAX_CHECKPOINT_AGE_MS || age < -60_000) {
        await fs.rm(file, { force: true }).catch(() => {});
        return null;
      }
      return parsed;
    } catch (error) {
      if (error?.code !== "ENOENT") {
        await fs.rm(file, { force: true }).catch(() => {});
      }
      return null;
    }
  }

  async function summary(origin, conversationId, options = {}) {
    const data = await load(origin, conversationId, options);
    return data ? summarizeCheckpoint(data) : null;
  }

  async function clear(origin, conversationId) {
    if (!origin || !conversationId) return false;
    const file = filePath(origin, conversationId);
    return enqueue(file, async () => {
      await fs.rm(file, { force: true }).catch(() => {});
      return true;
    });
  }

  async function clearForOrigin(origin) {
    if (!origin) return;
    await fs.rm(originDir(origin), { recursive: true, force: true }).catch(() => {});
  }

  async function clearAll() {
    await fs.rm(baseDir, { recursive: true, force: true }).catch(() => {});
  }

  return {
    save,
    load,
    summary,
    clear,
    clearForOrigin,
    clearAll
  };
}

module.exports = {
  CHECKPOINT_VERSION,
  MAX_CHECKPOINT_AGE_MS,
  STAGE_LABELS,
  createCheckpointStore,
  describeStage,
  summarizeCheckpoint
};
