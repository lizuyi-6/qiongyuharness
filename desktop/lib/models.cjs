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

function modelServiceError(status, detail) {
  if (status === 451) return "当前网络无法访问模型服务(451)";
  if (status === 401) return "模型凭证无效，请重新登录后重试";
  if (typeof detail === "string" && detail.trim()) return detail.trim();
  return "模型列表加载失败，请检查网络连接后重试";
}

async function listModelsWithRefresh(getCredential, fetchModels) {
  let credential = await getCredential(false);
  try {
    return await fetchModels(credential);
  } catch (error) {
    if (error.status !== 401) throw error;
    credential = await getCredential(true);
    return fetchModels(credential);
  }
}

module.exports = { listModelsWithRefresh, modelServiceError, parseModels };
