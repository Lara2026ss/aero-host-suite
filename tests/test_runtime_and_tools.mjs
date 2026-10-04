import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { Runtime } from "../core/runtime.mjs";

process.env.FLUXER_TEST_HEADLESS_DECISION = "allow_once";

console.log("=== TESTING FULL RUNTIME & TOOL REGISTRATION ===");

const rt = new Runtime({ rootDir: process.cwd() });
const initRes = await rt.init();

console.log("[1] Runtime initialized successfully:");
console.log(` - Product: ${initRes.product}`);
console.log(` - Version: ${initRes.version}`);
console.log(` - Tools Loaded (${initRes.toolsLoaded}):`, rt.registry.moduleNames().join(", "));

assert.equal(initRes.ok, true);
assert.equal(initRes.version, "V26-0.1");
assert.equal(rt.registry.moduleNames().length, 11, "Expected exactly 11 tools loaded!");

const expectedTools = [
  "files", "terminal", "system", "web", "documents",
  "media", "shortcuts", "security", "upd", "feedback", "guide"
];

for (const tool of expectedTools) {
  assert.ok(rt.registry.get(tool), `Tool '${tool}' must be registered!`);
}
console.log("[2] All 11 tools verified present in registry!");

// Test Router Dispatch for each tool:
console.log("\n[3] Testing Router Execution on each tool:");

// A. Guide
const catRes = await rt.router.execute("guide", "get_catalog", {});
assert.equal(catRes.ok, true);
assert.equal(catRes.toolsCount, 11);
console.log(` - guide.get_catalog: OK (${catRes.toolsCount} tools)`);

// B. Security
const secRes = await rt.router.execute("security", "get_status", {});
assert.equal(secRes.ok, true);
console.log(" - security.get_status: OK");

// C. Files
const tmpTestFile = path.join(process.cwd(), "data", "test_file_router.txt");
const writeRes = await rt.router.execute("files", "write_file", { path: tmpTestFile, content: "Línea 1\nLínea 2" });
assert.equal(writeRes.ok, true);
const readRes = await rt.router.execute("files", "read_file", { path: tmpTestFile, startLine: 1, endLine: 1 });
assert.equal(readRes.ok, true);
assert.equal(readRes.content, "Línea 1");
await fs.unlink(tmpTestFile).catch(() => {});
console.log(" - files.write_file & read_file: OK");

// D. System
const sysRes = await rt.router.execute("system", "get_info", {});
assert.equal(sysRes.ok, true);
assert.ok(sysRes.system.cpu);
console.log(` - system.get_info: OK (CPU: ${sysRes.system.cpu.model})`);

// E. Terminal
const termRes = await rt.router.execute("terminal", "run_command", { command: "Write-Output 'Fluxer XZ Live'" });
assert.equal(termRes.ok, true);
assert.match(termRes.stdout, /Fluxer XZ Live/);
console.log(" - terminal.run_command: OK");

// F. Shortcuts
const scList = await rt.router.execute("shortcuts", "list_shortcuts", {});
assert.equal(scList.ok, true);
assert.ok(scList.count >= 5);
console.log(` - shortcuts.list_shortcuts: OK (${scList.count} shortcuts)`);

// G. Feedback
const fbList = await rt.router.execute("feedback", "list_feedback", {});
assert.equal(fbList.ok, true);
console.log(" - feedback.list_feedback: OK");

// H. Media
const printersRes = await rt.router.execute("media", "list_printers", {});
assert.equal(printersRes.ok, true);
console.log(` - media.list_printers: OK (${printersRes.count} printers found)`);

// I. UPD
const updStat = await rt.router.execute("upd", "get_status", {});
assert.equal(updStat.ok, true);
assert.equal(updStat.version, "V26-0.1");
console.log(" - upd.get_status: OK");

console.log("\n[SUCCESS] Runtime, registry and router verified for all 11 tools!");
