const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { app, net, safeStorage } = require("electron");

const { createDesktopAgent } = require("../agent.cjs");
const fakeKey = "qo-smoke-fake-key-never-use-for-authentication";
const origin = "https://workplace.invalid";
let temporaryUserData;
let originalNetFetch;

async function credentialFiles(root) {
  const found = [];
  async function walk(directory) {
    let entries;
    try { entries = await fs.readdir(directory, { withFileTypes: true }); }
    catch (error) {
      if (error.code === "ENOENT") return;
      throw error;
    }
    for (const entry of entries) {
      const target = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(target);
      else if (entry.isFile()) found.push(target);
    }
  }
  await walk(root);
  return found;
}

async function main() {
  temporaryUserData = await fs.mkdtemp(path.join(os.tmpdir(), "qo-safe-storage-smoke-"));
  app.setPath("userData", temporaryUserData);
  await app.whenReady();

  const encryptionAvailable = safeStorage.isEncryptionAvailable();
  console.log(`safeStorage.isEncryptionAvailable()=${encryptionAvailable}`);
  assert.equal(encryptionAvailable, true, "safeStorage encryption is unavailable");

  let credentialRequestCount = 0;
  let rejectCredentialRequest = false;
  const session = {
    cookies: {
      async get() { return [{ name: "qo_session", value: "s".repeat(43) }]; }
    },
    async fetch(url, options = {}) {
      const parsed = new URL(url);
      if (parsed.pathname === "/api/v1/me" && options.method !== "POST") {
        return new Response(JSON.stringify({ id: "safe-storage-smoke-user" }), { status: 200 });
      }
      if (parsed.pathname === "/api/v1/me/model-credential" && options.method === "POST") {
        credentialRequestCount += 1;
        if (rejectCredentialRequest) return new Response("unavailable", { status: 503 });
        return new Response(JSON.stringify({
          api_key: fakeKey,
          key_id: 1,
          group_id: 1,
          base_url: "https://model.invalid/v1"
        }), { status: 200 });
      }
      throw new Error(`Unexpected session request: ${options.method || "GET"} ${url}`);
    }
  };
  const fakeWindow = {
    isDestroyed: () => false,
    webContents: { session }
  };
  const agent = createDesktopAgent({
    app: { getPath: () => temporaryUserData, isPackaged: false },
    getWindow: () => fakeWindow,
    getOrigin: () => origin
  });

  originalNetFetch = net.fetch;
  let observedAuthorization;
  net.fetch = async (_url, options = {}) => {
    observedAuthorization = new Headers(options.headers).get("Authorization");
    return new Response(JSON.stringify({ data: [{ id: "safe-storage-smoke-model" }] }), { status: 200 });
  };

  const models = await agent.prefetchAfterLogin();
  assert.deepEqual(models, ["safe-storage-smoke-model"]);
  assert.equal(observedAuthorization, `Bearer ${fakeKey}`, "stored key did not decrypt for production model lookup");

  const root = path.join(temporaryUserData, "qo-model-credentials");
  const files = await credentialFiles(root);
  assert.equal(files.length, 1, "expected one persisted credential file");
  for (const file of files) {
    const bytes = await fs.readFile(file);
    assert.equal(bytes.includes(Buffer.from(fakeKey, "utf8")), false, "credential file contains plaintext fake key");
  }
  console.log(`credentialReadBack=models=${JSON.stringify(models)} authorizationMatchesFakeKey=true`);
  console.log(`plaintextScan=files:${files.length} containsFakeKey:false`);

  await agent.clearCredentialsForOrigin(origin);
  const remainingAfterDelete = await credentialFiles(root);
  assert.deepEqual(remainingAfterDelete, [], "credential files remain after deleting origin credentials");
  rejectCredentialRequest = true;
  const missingRead = await agent.models();
  assert.deepEqual(missingRead.models, [], "credential lookup after deletion unexpectedly returned models");
  assert.equal(credentialRequestCount, 2, "lookup after deletion did not request a missing credential");
  assert.deepEqual(await credentialFiles(root), [], "failed credential refresh recreated a credential file");
  console.log(`deleteReadBack=models:${JSON.stringify(missingRead.models)} credentialRequests:${credentialRequestCount} files:0`);
  console.log("safeStorage smoke assertions passed");
}

main().catch((error) => {
  console.error(error?.stack || error);
  process.exitCode = 1;
}).finally(async () => {
  if (originalNetFetch) net.fetch = originalNetFetch;
  if (app.isReady()) app.quit();
  if (temporaryUserData) await fs.rm(temporaryUserData, { recursive: true, force: true });
});
