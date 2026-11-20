const { quotaExhaustionMessage } = require("./usage.cjs");

function parseModels(payload) {
  if (!payload || typeof payload !== "object" || !Array.isArray(payload.data)) {
    throw new Error("模型服务返回的列表格式无效");
  }
  const models = [];
  const seen = new Set();
  for (const item of payload.data) {
    if (!item || typeof item.id !== "string") continue;
    const id = item.id.trim();
    if (id && !seen.has(id)) {
      seen.add(id);
      models.push(id);
    }
  }
  return models;
}

function modelServiceCode(payload) {
  const candidates = [payload?.code, payload?.error?.code, payload?.error?.type];
  return candidates.find((value) => typeof value === "string") || null;
}

function modelServiceMessage(payload) {
  const candidates = [payload?.message, payload?.error?.message, payload?.detail];
  return candidates.find((value) => typeof value === "string" && value.trim())?.trim() || "";
}

function modelServiceError(status, payload) {
  if (status === 451) return "当前网络无法访问模型服务(451)";
  if (status === 401) return "模型凭证无效，请重新登录后重试";
  if (status === 429) {
    const quotaMessage = quotaExhaustionMessage({
      type: payload?.type || payload?.error?.type,
      code: modelServiceCode(payload),
      message: modelServiceMessage(payload),
      resetAt: payload?.resetAt || payload?.error?.resetAt,
      retryAfter: payload?.retryAfter || payload?.error?.retryAfter
    }, null);
    return quotaMessage || "模型额度已用完，请稍后再试";
  }
  if (status === 403) {
    const code = modelServiceCode(payload);
    if (code === "INSUFFICIENT_BALANCE") return "模型账户余额不足，请联系管理员开通穹驭套餐";
    if (code === "SUBSCRIPTION_NOT_FOUND" || code === "SUBSCRIPTION_INVALID") {
      return "尚未开通穹驭套餐或套餐已失效，请联系管理员";
    }
    if (code === "API_KEY_EXPIRED") return "模型凭证已过期，请重新登录后重试";
    if (code === "ACCESS_DENIED") return "当前网络不允许访问模型服务(403)";
    if (["GROUP_DISABLED", "GROUP_DELETED", "GROUP_NOT_ALLOWED"].includes(code)) {
      return "穹驭分组当前不可用，请联系管理员";
    }
  }
  const base = `模型服务返回错误(HTTP ${status})`;
  const message = modelServiceMessage(payload).replace(/[\u0000-\u001f\u007f-\u009f]/gu, "").slice(0, 200);
  return message ? `${base}: ${message}` : base;
}

async function listModelsWithRefresh(getCredential, fetchModels) {
  let credential = await getCredential(false);
  try {
    return await fetchModels(credential);
  } catch (error) {
    const refreshable = error.status === 401 ||
      (error.status === 403 && ["SUBSCRIPTION_NOT_FOUND", "INSUFFICIENT_BALANCE"].includes(error.code));
    if (!refreshable) throw error;
    credential = await getCredential(true);
    return fetchModels(credential);
  }
}

module.exports = { listModelsWithRefresh, modelServiceCode, modelServiceError, parseModels };
