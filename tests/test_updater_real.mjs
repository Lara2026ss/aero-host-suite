import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import http from "node:http";
import crypto from "node:crypto";
import { execSync } from "node:child_process";
import { UpdaterEngine } from "../updater/updater.mjs";

console.log("=== CRITERION G: FULL REMOTE UPDATER PHYSICAL TEST ===");

const sandboxDir = path.join(process.cwd(), "data", "test_sandbox_install");
if (await fs.stat(sandboxDir).catch(() => null)) {
  await fs.rm(sandboxDir, { recursive: true, force: true });
}
await fs.mkdir(sandboxDir, { recursive: true });

// 1. Initial clean install of V26-0.1 in sandbox
await fs.writeFile(path.join(sandboxDir, "VERSION"), "V26-0.1\n", "utf8");
await fs.writeFile(path.join(sandboxDir, "package.json"), JSON.stringify({ name: "fluxer-xz", version: "26.0.1" }), "utf8");
await fs.writeFile(path.join(sandboxDir, "user_config.json"), JSON.stringify({ customSettings: true }), "utf8");

// 1. Prepare physical test archive for V26-0.2
const updateSourceDir = path.join(process.cwd(), "data", "test_v26_0_2_src");
await fs.rm(updateSourceDir, { recursive: true, force: true });
await fs.mkdir(updateSourceDir, { recursive: true });
await fs.writeFile(path.join(updateSourceDir, "VERSION"), "V26-0.2\n", "utf8");
await fs.writeFile(path.join(updateSourceDir, "package.json"), JSON.stringify({ name: "fluxer-xz", version: "26.0.2" }), "utf8");
await fs.writeFile(path.join(updateSourceDir, "update_marker.txt"), "Update to V26-0.2 verified physical\n", "utf8");

const testZipPath = path.join(process.cwd(), "data", "test_v26_0_2.zip");
await fs.rm(testZipPath, { force: true });
execSync(`tar.exe -a -c -f "${testZipPath}" *`, { cwd: updateSourceDir });

const testZipBuffer = await fs.readFile(testZipPath);
const correctSha256 = crypto.createHash("sha256").update(testZipBuffer).digest("hex").toUpperCase();

// Start local HTTP server to simulate GitHub release endpoint
const server = http.createServer((req, res) => {
  if (req.url.includes("releases")) {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({
      tag_name: "v26.0.2",
      name: "FLUXER XZ - PUBLIC RELEASE V26-0.2",
      assets: [
        {
          name: "fluxer-v26.0.2.zip",
          browser_download_url: `http://127.0.0.1:${server.address().port}/download/fluxer-v26.0.2.zip`
        }
      ]
    }));
  } else if (req.url.includes("download")) {
    res.writeHead(200, { "Content-Type": "application/zip", "Content-Length": testZipBuffer.length });
    res.end(testZipBuffer);
  } else {
    res.writeHead(404);
    res.end();
  }
});

await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const serverPort = server.address().port;
const serverBase = `http://127.0.0.1:${serverPort}`;

const updater = new UpdaterEngine({
  rootDir: sandboxDir,
  apiBase: serverBase,
  repoOwner: "test-owner",
  repoName: "test-repo"
});

// 2. Check update against HTTP server
const checkRes = await updater.check();
console.log("[1] Release Check Result:", checkRes);
assert.equal(checkRes.ok, true);
assert.match(checkRes.latestVersion, /^V26[-.]0\.?2$/i);
assert.equal(checkRes.hasUpdate, true);

const realPkgUrl = checkRes.packageUrl;

// 3. Test Checksum Mismatch Security Rejection
console.log("\n[2] Testing SHA-256 Mismatch Rejection...");
await assert.rejects(async () => {
  await updater.apply({
    packageUrl: realPkgUrl,
    expectedSha256: "0000000000000000000000000000000000000000000000000000000000000000",
    force: true,
    releaseOverride: checkRes
  });
}, /CHECKSUM_MISMATCH/, "Updater MUST reject corrupted/tampered package!");
console.log("[PASS] Checksum mismatch was successfully detected and aborted!");

// 4. Test Legitimate Apply with Correct Checksum
console.log("\n[3] Applying legitimate V26-0.2 update from GitHub...");
const applyRes = await updater.apply({
  packageUrl: realPkgUrl,
  expectedSha256: correctSha256,
  force: true,
  releaseOverride: checkRes
});
assert.equal(applyRes.ok, true);
assert.equal(applyRes.applied, true);
assert.match(applyRes.newVersion, /^V26[-.]0\.?2$/i);

// Physical disk verification
const diskVersion = (await fs.readFile(path.join(sandboxDir, "VERSION"), "utf8")).trim();
assert.equal(diskVersion, "V26-0.2", "Physical VERSION file on disk must be V26-0.2!");
const markerExists = await fs.readFile(path.join(sandboxDir, "update_marker.txt"), "utf8");
assert.match(markerExists, /Update to V26-0.2 verified/);
// Verify user config was NOT touched
const userConfig = JSON.parse(await fs.readFile(path.join(sandboxDir, "user_config.json"), "utf8"));
assert.equal(userConfig.customSettings, true, "User config was preserved!");
console.log("[PASS] Physical disk check verified: Version is V26-0.2 and update_marker is present!");

// 5. Test Rollback
console.log("\n[4] Testing Rollback...");
const rollbackRes = await updater.rollback({ backupPath: applyRes.backupPath });
assert.equal(rollbackRes.ok, true);
assert.equal(rollbackRes.restoredVersion, "V26-0.1");

// Physical disk check after rollback
const restoredDiskVersion = (await fs.readFile(path.join(sandboxDir, "VERSION"), "utf8")).trim();
assert.equal(restoredDiskVersion, "V26-0.1", "Physical VERSION file on disk must be reverted to V26-0.1!");
console.log("[PASS] Rollback physically verified on disk: Version reverted to V26-0.1!");

// Clean up sandbox and test files
await fs.rm(sandboxDir, { recursive: true, force: true });
await fs.rm(updateSourceDir, { recursive: true, force: true });
await fs.rm(testZipPath, { force: true });
server.close();

console.log("\n[SUCCESS] Criterion G: Remote update, checksum validation, disk check & rollback 100% verified!");
