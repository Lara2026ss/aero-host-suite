import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { SecurityManager, RISK_LEVELS } from "../core/permissions.mjs";

console.log("=== TESTING SECURITY & PERMISSIONS ENGINE ===");

const testManager = new SecurityManager({ rootDir: process.cwd() });

// 1. Low risk: Always allowed
const lowRes = await testManager.checkPermission({
  tool: "files",
  action: "read_file",
  risk: RISK_LEVELS.LOW,
  resource: "test.txt"
});
assert.equal(lowRes.allowed, true);
console.log("[1] Low risk auto-permission passed!");

// 2. Blocked path integrity test
const blockedRes = await testManager.checkPermission({
  tool: "files",
  action: "delete_path",
  risk: RISK_LEVELS.DESTRUCTIVE,
  resource: "C:\\Windows\\System32\\drivers\\etc\\hosts"
});
assert.equal(blockedRes.allowed, false);
assert.match(blockedRes.reason, /protegido por la integridad del sistema operativo/);
console.log("[2] System protected path blocking passed!");

// 3. Elevation behavior: High risk allowed when elevated
testManager.setElevated(true, "*");
const elevatedHighRes = await testManager.checkPermission({
  tool: "terminal",
  action: "run_command",
  risk: RISK_LEVELS.HIGH,
  resource: "Get-Process"
});
assert.equal(elevatedHighRes.allowed, true);
console.log("[3] Elevation allows HIGH risk actions consistently!");

// 4. Destructive action ALWAYS prompts even if elevated
// Simulate headless decision 'deny'
process.env.FLUXER_TEST_HEADLESS_DECISION = "deny";
const destructiveRes = await testManager.checkPermission({
  tool: "files",
  action: "delete_path",
  risk: RISK_LEVELS.DESTRUCTIVE,
  resource: "C:\\temp\\junk.txt"
});
assert.equal(destructiveRes.allowed, false);
console.log("[4] Destructive actions require prompt even when elevated!");

// Simulate headless decision 'allow_once'
process.env.FLUXER_TEST_HEADLESS_DECISION = "allow_once";
const destructiveAllowed = await testManager.checkPermission({
  tool: "files",
  action: "delete_path",
  risk: RISK_LEVELS.DESTRUCTIVE,
  resource: "C:\\temp\\junk.txt"
});
assert.equal(destructiveAllowed.allowed, true);
console.log("[5] Destructive action allowed once when user confirms!");

// 5. Verify audit log exists on disk and has entries
const auditContent = await fs.readFile(testManager.auditLogPath, "utf8");
assert.match(auditContent, /elevation_change/);
assert.match(auditContent, /blocked_system_integrity/);
console.log("[6] Physical audit log verified on disk at:", testManager.auditLogPath);

console.log("\n[PASS] All security manager tests passed!");
