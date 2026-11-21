const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { createDesktopLogger, redactSecrets } = require("../lib/logger.cjs");

test("redactSecrets strips tokens, api keys, cookies and base64 images", () => {
  const raw = [
    "Authorization: Bearer eyJhbGciOiJIUzI1NiJ9.abcdef123456",
    "key=sk-proj-1234567890abcdef",
    "Cookie: qo_session=secret_session_value_12345; theme=dark",
    '{"api_key":"my-super-secret-val","password":"p@ssword123","ok":true}',
    "img=data:image/png;base64," + "A".repeat(120)
  ].join("\n");
  const safe = redactSecrets(raw);
  assert.doesNotMatch(safe, /eyJhbGciOi/);
  assert.doesNotMatch(safe, /sk-proj-123456/);
  assert.doesNotMatch(safe, /secret_session_value_12345/);
  assert.doesNotMatch(safe, /my-super-secret-val/);
  assert.doesNotMatch(safe, /p@ssword123/);
  assert.match(safe, /\[REDACTED\]/);
  assert.match(safe, /\[REDACTED_KEY\]/);
  assert.match(safe, /\[IMAGE_BASE64_OMITTED\]/);
});

test("createDesktopLogger persists logs and collects redacted diagnostics", async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "qo-logger-test-"));
  try {
    const logger = createDesktopLogger({ logDir: dir });
    logger.info("main", "startup ok", { version: "0.2.21" });
    logger.error("agent.run", "request failed with Bearer tok_secret_9999999", {
      api_key: "sk-hidden-value-999",
      stage: "execute"
    });
    logger.flushSync();

    const diag = await logger.collectDiagnostics({
      appVersion: "0.2.21",
      origin: "https://office.aethernettech.dev",
      updaterState: { status: "idle", latestVersion: "0.2.21" },
      checkpointSummary: { stage: "execute", stageLabel: "执行" }
    });
    assert.equal(diag.app_version, "0.2.21");
    assert.equal(diag.origin, "https://office.aethernettech.dev");
    assert.equal(diag.checkpoint_summary.stage, "execute");
    assert.match(diag.recent_logs, /startup ok/);
    assert.match(diag.recent_logs, /request failed with Bearer \[REDACTED\]/);
    assert.doesNotMatch(diag.recent_logs, /tok_secret_9999999/);
    assert.doesNotMatch(diag.recent_logs, /sk-hidden-value-999/);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
