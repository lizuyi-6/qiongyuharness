const fs = require("node:fs");
const fsp = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");

const MAX_MEMORY_ENTRIES = 400;
const MAX_FILE_BYTES = 1_500_000;
const DEFAULT_COLLECT_BYTES = 120_000;

function redactSecrets(input) {
  let text = String(input ?? "");
  if (!text) return "";
  // Authorization / Bearer tokens
  text = text.replace(/\b(Bearer\s+)[A-Za-z0-9._~+/=-]{8,}/gi, "$1[REDACTED]");
  // OpenAI / Sub2API / generic sk- keys
  text = text.replace(/\b(sk-[A-Za-z0-9_-]{8,})/g, "[REDACTED_KEY]");
  // Session cookies
  text = text.replace(/\b(qo_session|qo_browser)=([^;\s"'`]+)/gi, "$1=[REDACTED]");
  // JSON / key-value secret fields
  text = text.replace(
    /((?:["'])?(?:api[_-]?key|access[_-]?token|refresh[_-]?token|id[_-]?token|client[_-]?secret|authorization|cookie|set-cookie|password|secret|encryptedKey)(?:["'])?\s*[:=]\s*)(["']?)([^"',\s}\]]{4,})(\2)/gi,
    "$1$2[REDACTED]$4"
  );
  // Base64 data URIs (avoid huge image blobs in logs)
  text = text.replace(/data:image\/[a-zA-Z0-9.+-]+;base64,[A-Za-z0-9+/=]{64,}/g, "[IMAGE_BASE64_OMITTED]");
  return text;
}

function safeSerializeMeta(meta) {
  if (meta === undefined || meta === null) return "";
  if (typeof meta === "string") return redactSecrets(meta).slice(0, 1200);
  if (meta instanceof Error) {
    return redactSecrets(`${meta.name || "Error"}: ${meta.message || ""}${meta.stack ? `\n${meta.stack}` : ""}`).slice(0, 2000);
  }
  try {
    const seen = new WeakSet();
    const raw = JSON.stringify(meta, (key, value) => {
      if (/(?:api[_-]?key|token|secret|password|cookie|authorization|encryptedKey|png_base64|b64_json)/i.test(key)) {
        return "[REDACTED]";
      }
      if (typeof value === "object" && value !== null) {
        if (seen.has(value)) return "[Circular]";
        seen.add(value);
      }
      if (typeof value === "string" && value.length > 800) {
        return value.slice(0, 800) + `…(${value.length} chars)`;
      }
      return value;
    });
    return redactSecrets(raw).slice(0, 1800);
  } catch {
    return redactSecrets(String(meta)).slice(0, 800);
  }
}

function createDesktopLogger({ logDir, now = Date.now } = {}) {
  const entries = [];
  const logFile = logDir ? path.join(logDir, "app.log") : "";
  const prevFile = logDir ? path.join(logDir, "app.prev.log") : "";
  let dirReady = false;
  let pendingLines = [];
  let flushTimer = null;

  function ensureDirSync() {
    if (!logDir || dirReady) return Boolean(logDir);
    try {
      fs.mkdirSync(logDir, { recursive: true });
      dirReady = true;
      return true;
    } catch {
      return false;
    }
  }

  function flushSync() {
    if (!logFile || !pendingLines.length) return;
    if (!ensureDirSync()) {
      pendingLines = [];
      return;
    }
    const batch = pendingLines.join("");
    pendingLines = [];
    try {
      try {
        const stat = fs.statSync(logFile);
        if (stat.size >= MAX_FILE_BYTES) {
          try { fs.rmSync(prevFile, { force: true }); } catch {}
          fs.renameSync(logFile, prevFile);
        }
      } catch {}
      fs.appendFileSync(logFile, batch, { encoding: "utf8", mode: 0o600 });
    } catch {}
  }

  function scheduleFlush() {
    if (!logFile || flushTimer) return;
    flushTimer = setTimeout(() => {
      flushTimer = null;
      flushSync();
    }, 300);
    flushTimer.unref?.();
  }

  function log(level, scope, message, meta) {
    const ts = new Date(now()).toISOString();
    const lv = String(level || "INFO").toUpperCase();
    const sc = String(scope || "app").slice(0, 32);
    const msg = redactSecrets(message || "").replace(/\r?\n/g, " ").slice(0, 1200);
    const extra = meta !== undefined ? safeSerializeMeta(meta) : "";
    const line = `[${ts}] [${lv}] [${sc}] ${msg}${extra ? ` | ${extra}` : ""}`;
    entries.push(line);
    if (entries.length > MAX_MEMORY_ENTRIES) {
      entries.splice(0, entries.length - MAX_MEMORY_ENTRIES);
    }
    if (logFile) {
      pendingLines.push(line + "\n");
      if (lv === "ERROR" || lv === "WARN" || pendingLines.length >= 16) {
        flushSync();
      } else {
        scheduleFlush();
      }
    }
    return line;
  }

  async function readTail(filePath, maxBytes) {
    try {
      const stat = await fsp.stat(filePath);
      if (!stat.size) return "";
      const start = Math.max(0, stat.size - maxBytes);
      const handle = await fsp.open(filePath, "r");
      try {
        const length = stat.size - start;
        const buf = Buffer.alloc(length);
        const { bytesRead } = await handle.read(buf, 0, length, start);
        let text = buf.subarray(0, bytesRead).toString("utf8");
        if (start > 0) {
          const firstNl = text.indexOf("\n");
          if (firstNl >= 0) text = text.slice(firstNl + 1);
        }
        return text;
      } finally {
        await handle.close();
      }
    } catch {
      return "";
    }
  }

  async function readRecentLogText(maxBytes = DEFAULT_COLLECT_BYTES) {
    flushSync();
    let text = "";
    if (logFile) {
      text = await readTail(logFile, maxBytes);
      if (text.length < maxBytes * 0.5 && prevFile) {
        const older = await readTail(prevFile, Math.max(8_000, maxBytes - text.length));
        if (older) text = `${older.trimEnd()}\n${text}`.trim();
      }
    }
    if (!text && entries.length) {
      text = entries.join("\n");
    }
    const safe = redactSecrets(text);
    if (safe.length > maxBytes) {
      return safe.slice(safe.length - maxBytes);
    }
    return safe;
  }

  async function collectDiagnostics({
    appVersion = "",
    origin = "",
    updaterState = null,
    checkpointSummary = null,
    maxBytes = DEFAULT_COLLECT_BYTES
  } = {}) {
    const recentLogs = await readRecentLogText(maxBytes);
    const safeUpdater = updaterState && typeof updaterState === "object"
      ? {
          status: updaterState.status || "",
          latestVersion: updaterState.latestVersion || "",
          mode: updaterState.mode || "",
          error: redactSecrets(updaterState.error || "")
        }
      : null;
    return {
      collected_at: new Date(now()).toISOString(),
      app_version: String(appVersion || ""),
      electron_version: String(process.versions?.electron || ""),
      node_version: String(process.versions?.node || ""),
      platform: `${process.platform}/${process.arch}`,
      os_release: `${os.type()} ${os.release()}`,
      origin: String(origin || ""),
      updater_state: safeUpdater,
      checkpoint_summary: checkpointSummary || null,
      recent_logs: recentLogs
    };
  }

  return {
    log,
    info: (scope, message, meta) => log("INFO", scope, message, meta),
    warn: (scope, message, meta) => log("WARN", scope, message, meta),
    error: (scope, message, meta) => log("ERROR", scope, message, meta),
    flushSync,
    readRecentLogText,
    collectDiagnostics,
    getEntries: () => [...entries],
    logFile
  };
}

module.exports = {
  createDesktopLogger,
  redactSecrets
};
