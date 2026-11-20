const { randomUUID } = require("node:crypto");
const { spawn } = require("node:child_process");
const { StringDecoder } = require("node:string_decoder");
const { RESPONSE_LIMIT } = require("./loopback-gateway.cjs");
const { buildDshArgs } = require("./runtime-env.cjs");

function textFromMessage(event) {
  const message = event?.data?.message || event?.data || {};
  const blocks = Array.isArray(message.content) ? message.content : [];
  return blocks.filter((part) => part?.type === "text" && typeof part.text === "string")
    .map((part) => part.text).join("");
}

function abortError() {
  return Object.assign(new Error("已停止生成"), { code: "ABORTED" });
}

function runDsh({
  executable,
  dshBin,
  patchPath,
  contextPatchPath,
  extraPatchPaths = [],
  cwd,
  env,
  model,
  prompt,
  onEvent = () => {},
  timeoutMs = 620_000,
  signal,
  spawnImpl = spawn
}) {
  if (!executable || !model || typeof prompt !== "string" || !cwd || !env) {
    return Promise.reject(new Error("桌面运行参数无效"));
  }
  if (signal?.aborted) {
    return Promise.reject(abortError());
  }

  const args = buildDshArgs({ dshBin, patchPath, contextPatchPath, extraPatchPaths });
  const runId = randomUUID();
  let child;
  try {
    child = spawnImpl(executable, args, {
      cwd,
      env,
      windowsHide: true,
      stdio: ["pipe", "pipe", "ignore"]
    });
  } catch {
    return Promise.reject(new Error("无法启动桌面 Harness"));
  }

  let settled = false;
  let active = false;
  let promptSent = false;
  let finalText = "";
  let pending = "";
  const decoder = new StringDecoder("utf8");
  let processClosed = false;
  let resolveClose;
  const closed = new Promise((resolve) => { resolveClose = resolve; });
  let resolveResult;
  let rejectResult;
  const result = new Promise((resolve, reject) => { resolveResult = resolve; rejectResult = reject; });

  const fail = (error) => {
    if (settled) return;
    settled = true;
    rejectResult(error instanceof Error ? error : new Error("DSH 运行失败"));
  };
  const succeed = () => {
    if (settled) return;
    if (!finalText) {
      fail(new Error("DSH 未返回文本结果"));
      return;
    }
    settled = true;
    resolveResult(finalText);
  };
  const send = (message) => {
    if (!child.stdin.writable) throw new Error("DSH 输入通道已关闭");
    child.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", ...message })}\n`, "utf8");
  };

  async function handleLine(line) {
    if (!line.trim() || settled) return;
    let message;
    try { message = JSON.parse(line); } catch {
      throw new Error("DSH 返回了无效的协议消息");
    }
    if (message && Object.hasOwn(message, "id") && typeof message.method === "string") {
      send({ id: message.id, error: { code: -32601, message: "Policy denies tool and remote requests" } });
    }
    if (message?.error) throw new Error("DSH 协议拒绝了本次请求");
    if (message?.id === 1 && !promptSent) {
      promptSent = true;
      send({ id: 2, method: "session/prompt", params: {
        sessionId: runId,
        contentBlocks: [{ type: "text", text: prompt }]
      } });
    }

    const params = message?.params || {};
    if (params.sessionId !== runId) return;
    const event = params.event || {};
    if (event.type === "agent/inbox/spliced") active = true;
    if (event.type === "assistant/message") finalText = textFromMessage(event) || finalText;
    if (message.method === "session.status" && params.status === "idle" && active) succeed();
  }

  let lineQueue = Promise.resolve();
  child.stdout.on("data", (chunk) => {
    if (settled) return;
    pending += decoder.write(chunk);
    if (Buffer.byteLength(pending) > RESPONSE_LIMIT) {
      fail(new Error("DSH 协议输出超过大小限制"));
      return;
    }
    let index;
    while ((index = pending.indexOf("\n")) >= 0) {
      const line = pending.slice(0, index).replace(/\r$/, "");
      pending = pending.slice(index + 1);
      lineQueue = lineQueue.then(() => handleLine(line)).catch(fail);
    }
  });
  child.stdout.on("end", () => {
    pending += decoder.end();
    if (pending.trim() && !settled) lineQueue = lineQueue.then(() => handleLine(pending)).catch(fail);
  });
  child.once("error", () => fail(new Error("无法启动桌面 Harness")));
  child.once("close", (code) => {
    processClosed = true;
    resolveClose();
    if (!settled) fail(new Error(code === 0 ? "DSH 进程退出但没有完成响应" : "DSH 进程异常退出"));
  });

  try {
    send({ id: 1, method: "initialize", params: {
      cwd,
      provider: "company",
      model,
      approvalRequests: true
    } });
  } catch {
    fail(new Error("无法向桌面 Harness 发送初始化请求"));
  }

  const onAbort = () => fail(abortError());
  signal?.addEventListener?.("abort", onAbort, { once: true });

  const timer = setTimeout(() => fail(new Error("桌面 Harness 调用超时，请重试")), timeoutMs);
  timer.unref?.();
  return result.then((text) => text).finally(async () => {
    clearTimeout(timer);
    signal?.removeEventListener?.("abort", onAbort);
    if (!processClosed) {
      try { child.kill(); } catch {}
      const killTimer = new Promise((resolve) => setTimeout(resolve, 2_000));
      await Promise.race([closed, killTimer]);
      if (!processClosed) {
        try { child.kill("SIGKILL"); } catch {}
        await closed;
      }
    }
  });
}

module.exports = { runDsh, textFromMessage };
