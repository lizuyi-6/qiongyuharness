const test = require("node:test");
const assert = require("node:assert/strict");
const { runDsh } = require("../lib/dsh-bridge.cjs");
const { buildDshArgs, buildRuntimeEnv } = require("../lib/runtime-env.cjs");

test("runtime env 的旧 context 调用保留环境键与参数顺序", () => {
  const env = buildRuntimeEnv({ baseEnv: { PATH: "fake-path", SECRET: "do-not-copy" }, home: "C:/temp/home",
    gatewayUrl: "http://127.0.0.1:1234/v1", gatewayKey: "loopback-key", model: "m", system: "s",
    contextPath: "C:/temp/context.json", dshBin: "C:/runtime/dsh/bin.js" });
  assert.deepEqual(Object.keys(env).sort(), ["DSH_AGENTS_HOME", "DSH_HOME", "DSH_SYSTEM_PROMPT", "HOME",
    "NO_PROXY", "PATH", "QO_DSH_EXECUTABLE", "QO_DSH_PLUGIN_CONTEXT", "QO_RUNTIME_KEY", "QO_RUNTIME_MODEL",
    "QO_RUNTIME_URL", "USERPROFILE"].sort());
  assert.equal(env.QO_DSH_PLUGIN_CONTEXT, "C:/temp/context.json");
  assert.equal(env.QO_DSH_EXECUTABLE, "C:/runtime/dsh/bin.js");
  assert.equal(Object.hasOwn(env, "SECRET"), false);
  const oldArgs = buildDshArgs({ dshBin: "dsh.js", patchPath: "office.patch.yml", contextPatchPath: "context.patch.yml" });
  assert.deepEqual(oldArgs, ["dsh.js", "--profile", "sdk-minimal", "--patch", "office.patch.yml",
    "--patch", "context.patch.yml"]);
});

test("runtime 支持 context 和 image 多 patch，且 broker key 只在 image 启用时透传", () => {
  const args = buildDshArgs({ dshBin: "dsh.js", patchPath: "office.patch.yml", contextPatchPath: "context.patch.yml",
    extraPatchPaths: ["image.patch.yml", "policy.patch.yml"] });
  assert.deepEqual(args, ["dsh.js", "--profile", "sdk-minimal", "--patch", "office.patch.yml",
    "--patch", "context.patch.yml", "--patch", "image.patch.yml", "--patch", "policy.patch.yml"]);
  const env = buildRuntimeEnv({ baseEnv: {}, home: "C:/temp", gatewayUrl: "http://127.0.0.1:1234/v1",
    gatewayKey: "local", model: "m", system: "s", imageBrokerUrl: "http://127.0.0.1:8080",
    imageBrokerKey: "one-time-key", dshBin: "C:/runtime/dsh.js" });
  assert.equal(env.QO_IMAGE_BROKER_URL, "http://127.0.0.1:8080");
  assert.equal(env.QO_IMAGE_BROKER_KEY, "one-time-key");
  assert.equal(env.QO_DSH_EXECUTABLE, "C:/runtime/dsh.js");
  assert.equal(Object.hasOwn(env, "QO_DSH_PLUGIN_CONTEXT"), false);
});

test("dsh-bridge 旧调用与 image 多 patch 调用都把参数数组原样传给子进程", async () => {
  async function capturedArgs(options = {}) {
    let invocation;
    await assert.rejects(runDsh({ executable: "node.exe", dshBin: "dsh.js", patchPath: "office.patch.yml",
      cwd: "C:/temp", env: {}, model: "m", prompt: "p", ...options,
      spawnImpl: (executable, args) => {
        invocation = { executable, args };
        throw new Error("test stops before process creation");
      } }), /无法启动桌面 Harness/);
    return invocation;
  }

  assert.deepEqual(await capturedArgs({ contextPatchPath: "context.patch.yml" }), {
    executable: "node.exe",
    args: ["dsh.js", "--profile", "sdk-minimal", "--patch", "office.patch.yml", "--patch", "context.patch.yml"]
  });
  assert.deepEqual(await capturedArgs({ contextPatchPath: "context.patch.yml",
    extraPatchPaths: ["image.patch.yml", "policy.patch.yml"] }), {
    executable: "node.exe",
    args: ["dsh.js", "--profile", "sdk-minimal", "--patch", "office.patch.yml", "--patch", "context.patch.yml",
      "--patch", "image.patch.yml", "--patch", "policy.patch.yml"]
  });
});
