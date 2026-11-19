const path = require("node:path");

const PASSTHROUGH_ENV = ["PATH", "SystemRoot", "WINDIR", "TEMP", "TMP", "LANG"];

function buildRuntimeEnv({ baseEnv = process.env, home, gatewayUrl, gatewayKey, model, system, contextPath, dshBin }) {
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
  if (contextPath) {
    if (!dshBin) throw new Error("缺少 DSH 执行路径");
    env.QO_DSH_PLUGIN_CONTEXT = contextPath;
    env.QO_DSH_EXECUTABLE = dshBin;
  }
  return env;
}

function buildDshArgs({ dshBin, patchPath, contextPatchPath }) {
  if (!dshBin || !patchPath) throw new Error("缺少 Harness 运行文件");
  const args = [dshBin, "--profile", "sdk-minimal", "--patch", patchPath];
  if (contextPatchPath) args.push("--patch", contextPatchPath);
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
