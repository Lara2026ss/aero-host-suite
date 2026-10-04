import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { createFeedbackDomain } from "../tools/feedback.mjs";

console.log("=== TESTING FEEDBACK ENGINE (CRITERION H) ===");

const testFolder = path.join(process.cwd(), "data", "test_feedback");
await fs.mkdir(testFolder, { recursive: true });

const domain = createFeedbackDomain({
  runtime: { dirs: { feedback: testFolder } }
});

// 1. Send feedback with personal path that MUST be sanitized
const sendRes = await domain.actions.send_feedback({
  type: "bug",
  tool: "files",
  action: "patch_file",
  description: "Error al aplicar parche en archivo con salto de línea CRLF",
  severity: "high",
  context: {
    filePath: "C:\\Users\\mauri\\Documents\\secret_project\\app.js",
    errorDetail: "Range invalid"
  }
});

assert.equal(sendRes.ok, true);
assert.ok(sendRes.id);
assert.match(sendRes.githubIssueUrl, /https:\/\/github\.com\/Lara2026ss\/aero-fluxer-x\/issues\/new/);
console.log("[1] Feedback sent and GitHub URL generated:", sendRes.id);

// Check that personal user path was sanitized
const viewRes = await domain.actions.view_feedback({ id: sendRes.id });
assert.equal(viewRes.ok, true);
assert.equal(viewRes.feedback.context.filePath.includes("mauri"), false, "Personal path was not sanitized!");
assert.match(viewRes.feedback.context.filePath, /~/);
console.log("[2] Privacy sanitation verified: path was sanitized to ~");

// 2. List feedback
const listRes = await domain.actions.list_feedback({});
assert.equal(listRes.ok, true);
assert.ok(listRes.count >= 1);
console.log("[3] List feedback verified:", listRes.count, "items");

// 3. Update status
const updateRes = await domain.actions.update_status({ id: sendRes.id, status: "resuelto" });
assert.equal(updateRes.ok, true);
assert.equal(updateRes.feedback.status, "resuelto");
console.log("[4] Update status verified to 'resuelto'!");

// 4. Export feedback
const exportRes = await domain.actions.export_feedback({ id: sendRes.id });
assert.equal(exportRes.ok, true);
assert.equal(exportRes.data.id, sendRes.id);
console.log("[5] Export feedback verified!");

// 5. Delete feedback
const delRes = await domain.actions.delete_feedback({ id: sendRes.id });
assert.equal(delRes.ok, true);
const listAfterDel = await domain.actions.list_feedback({});
assert.equal(listAfterDel.feedbacks.some(f => f.id === sendRes.id), false);
console.log("[6] Delete feedback verified!");

// Cleanup
await fs.rm(testFolder, { recursive: true, force: true });
console.log("\n[PASS] All feedback tests passed successfully!");
