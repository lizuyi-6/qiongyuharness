const path = require("node:path");

const PASSTHROUGH_ENV = ["PATH", "SystemRoot", "WINDIR", "TEMP", "TMP", "LANG"];

function buildRuntimeEnv({ baseEnv = process.env, home, gatewayUrl, gatewayKey, model, system, contextPath,
  imageBrokerUrl, imageBrokerKey, searchBrokerUrl, searchBrokerKey, serverTime, dshBin }) {
  if (!home || !gatewayUrl || !gatewayKey || !model) throw new Error("桌面运行环境参数无效");
  const env = {};
  for (const name of PASSTHROUGH_ENV) {
    if (typeof baseEnv[name] === "string") env[name] = baseEnv[name];
  }
  env.HOME = home;
  env.USERPROFILE = home;
  env.DSH_HOME = home;
  env.DSH_AGENTS_HOME = home;
  env.NO_PROXY = "127.0.0.1,localhost";
  env.QO_RUNTIME_URL = gatewayUrl;
  env.QO_RUNTIME_KEY = gatewayKey;
  env.QO_RUNTIME_MODEL = model;
  env.DSH_SYSTEM_PROMPT = typeof system === "string" ? system : "";
  if (contextPath || imageBrokerUrl || imageBrokerKey || searchBrokerUrl || searchBrokerKey) {
    if (!dshBin) throw new Error("缺少 DSH 执行路径");
  }
  if (contextPath) {
    env.QO_DSH_PLUGIN_CONTEXT = contextPath;
  }
  if (contextPath || imageBrokerUrl || imageBrokerKey || searchBrokerUrl || searchBrokerKey) {
    env.QO_DSH_EXECUTABLE = dshBin;
  }
  if (imageBrokerUrl || imageBrokerKey) {
    if (!imageBrokerUrl || !imageBrokerKey) throw new Error("图片服务凭证无效");
    env.QO_IMAGE_BROKER_URL = imageBrokerUrl;
    env.QO_IMAGE_BROKER_KEY = imageBrokerKey;
  }
  if (searchBrokerUrl || searchBrokerKey) {
    if (!searchBrokerUrl || !searchBrokerKey) throw new Error("联网搜索服务凭证无效");
    env.QO_SEARCH_BROKER_URL = searchBrokerUrl;
    env.QO_SEARCH_BROKER_KEY = searchBrokerKey;
  }
  if (serverTime && typeof serverTime === "object") {
    if (serverTime.unix_timestamp) env.QO_SERVER_TIMESTAMP = String(serverTime.unix_timestamp);
    if (serverTime.year) env.QO_SERVER_YEAR = String(serverTime.year);
    if (serverTime.date) env.QO_SERVER_DATE = String(serverTime.date);
    if (serverTime.display) env.QO_SERVER_TIME_DISPLAY = String(serverTime.display);
  }
  return env;
}

function buildDshArgs({ dshBin, patchPath, contextPatchPath, extraPatchPaths = [] }) {
  if (!dshBin || !patchPath) throw new Error("缺少 Harness 运行文件");
  if (!Array.isArray(extraPatchPaths) || extraPatchPaths.some((patch) => typeof patch !== "string" || !patch)) {
    throw new Error("Harness 扩展 patch 参数无效");
  }
  const args = [dshBin, "--profile", "sdk-minimal", "--patch", patchPath];
  if (contextPatchPath) args.push("--patch", contextPatchPath);
  for (const patch of extraPatchPaths) args.push("--patch", patch);
  return args;
}

function resolveRuntimePaths(resourcesPath, appDirectory) {
  const runtimeRoot = resourcesPath
    ? path.join(resourcesPath, "qo-runtime")
    : path.resolve(appDirectory, "..");
  const harnessRoot = path.join(runtimeRoot, "harness");
  return {
    runtimeRoot,
    harnessRoot,
    patchPath: path.join(harnessRoot, "office.patch.yml"),
    contextPlugin: path.join(harnessRoot, "office-context.mjs"),
    dshBin: path.join(harnessRoot, "node_modules", "@deepseek-ai", "dsh", "lib", "bin.js"),
    bundledNode: path.join(runtimeRoot, "node", "node.exe")
  };
}

module.exports = { PASSTHROUGH_ENV, buildDshArgs, buildRuntimeEnv, resolveRuntimePaths };
