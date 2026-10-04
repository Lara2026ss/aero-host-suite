import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { createShortcutsDomain } from "../tools/shortcuts.mjs";

console.log("=== TESTING SHORTCUTS ENGINE (TEMPLATES, DRY-RUN & ROLLBACK) ===");

const testFolder = path.join(process.cwd(), "data", "test_shortcuts");
await fs.mkdir(testFolder, { recursive: true });

const domain = createShortcutsDomain({
  runtime: { dirs: { shortcuts: testFolder } }
});

// Wait briefly for default seed
await new Promise(r => setTimeout(r, 200));

// 1. List shortcuts
const listRes = await domain.actions.list_shortcuts();
assert.equal(listRes.ok, true);
console.log(`[1] Shortcuts listed: ${listRes.count} found in ${listRes.folder}`);
assert.ok(listRes.count >= 5, "Expected at least 5 default shortcuts seeded");

// 2. View shortcut
const viewRes = await domain.actions.view_shortcut({ id: "create_module_scaffold" });
assert.equal(viewRes.ok, true);
assert.equal(viewRes.shortcut.id, "create_module_scaffold");
console.log("[2] View shortcut passed:", viewRes.shortcut.name);

// 3. Dry-run execution with parameter interpolation
const dryRes = await domain.actions.run_shortcut({
  id: "create_module_scaffold",
  params: { name: "demo-auth", author: "Leo Tester" },
  dry_run: true
});
assert.equal(dryRes.ok, true);
assert.equal(dryRes.dryRun, true);
assert.equal(dryRes.plan.length, 2);
assert.equal(dryRes.plan[0].resolvedArgs.path, "src/demo-auth/index.js");
assert.match(dryRes.plan[0].resolvedArgs.content, /Módulo: demo-auth/);
assert.match(dryRes.plan[0].resolvedArgs.content, /Autor: Leo Tester/);
console.log("[3] Dry-run parameter interpolation verified:", dryRes.plan[0].resolvedArgs.path);

// 4. Create custom shortcut
const customRes = await domain.actions.create_shortcut({
  id: "test_custom_flow",
  name: "Custom Flow",
  description: "Test flow for rollback",
  params: { targetFile: "temp_test_file.txt" },
  steps: [
    {
      tool: "files",
      action: "write_file",
      args: { path: "{{targetFile}}", content: "hello world" },
      onError: "rollback"
    },
    {
      tool: "terminal",
      action: "run_command",
      args: { command: "FORCE_FAIL_COMMAND_123" },
      onError: "rollback"
    }
  ]
});
assert.equal(customRes.ok, true);
console.log("[4] Custom shortcut created successfully.");

// 5. Test rollback on failure
// Mock router that fails on step 2
const rollbackFile = path.join(process.cwd(), "temp_test_rollback.txt");
await fs.writeFile(rollbackFile, "should be deleted by rollback", "utf8");

const rollbackDomain = createShortcutsDomain({
  runtime: {
    dirs: { shortcuts: testFolder },
    router: {
      async execute(tool, action, args) {
        if (action === "run_command") throw new Error("Simulated terminal command failure");
        return { ok: true };
      }
    }
  }
});

const runFailRes = await rollbackDomain.actions.run_shortcut({
  id: "test_custom_flow",
  params: { targetFile: rollbackFile }
});
assert.equal(runFailRes.ok, false);
assert.equal(runFailRes.rollbackExecuted, true);
assert.equal(runFailRes.failedAtStep, 2);

// Check that rollbackFile was unlinked by rollback
let fileStillExists = true;
try {
  await fs.access(rollbackFile);
} catch {
  fileStillExists = false;
}
assert.equal(fileStillExists, false, "Rollback failed to delete the unlinked file!");
console.log("[5] Rollback execution verified! File successfully cleaned up after failure.");

// 6. Test set_folder
const newTestFolder = path.join(process.cwd(), "data", "test_shortcuts_alt");
const setFolderRes = await domain.actions.set_folder({ path: newTestFolder });
assert.equal(setFolderRes.ok, true);
console.log("[6] set_folder verified!");

// 7. Validate shortcut error checking
const invalidRes = await domain.actions.validate_shortcut({ shortcut: { id: "bad" } });
assert.equal(invalidRes.ok, false);
console.log("[7] Shortcut validation rejection verified:", invalidRes.error);

console.log("\n[PASS] All shortcuts engine unit tests passed successfully!");
