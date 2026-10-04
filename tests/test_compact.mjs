import assert from "node:assert/strict";
import { CompactEngine } from "../core/compact.mjs";

console.log("=== TESTING COMPACT ENGINE (ROUND-TRIP & TOKEN EFFICIENCY) ===");

const engine = new CompactEngine({ mode: "auto" });

// 1. Content Stash
const largeSnippet = "function calculateTotal(items) { return items.reduce((a, b) => a + b, 0); }";
const stashRes = engine.stashSet(largeSnippet);
console.log("[1] Stash created:", stashRes.stashId);
assert.equal(engine.stashHas(stashRes.stashId), true);
assert.equal(engine.stashGet(stashRes.stashId), largeSnippet);
assert.equal(engine.resolveContent(stashRes.stashId), largeSnippet);
assert.equal(engine.resolveContent("raw plain text"), "raw plain text");

// 2. Handles
const handle = engine.createHandle("file", "C:/Projects/demo.txt", { size: 1024 });
console.log("[2] Handle created:", handle);
assert.match(handle, /^f#\d+$/);
const resolved = engine.resolveHandle(handle);
assert.equal(resolved.type, "file");
assert.match(resolved.path, /demo\.txt/);
engine.releaseHandle(handle);
assert.equal(engine.resolveHandle(handle), null);

// 3. Smart read & Unchanged hash detection
const sampleContent = "Línea 1: Configuración inicial\nLínea 2: Parámetros del sistema\nLínea 3: Cierre";
const initialCheck = engine.checkUnchanged("test.txt", sampleContent);
assert.equal(initialCheck.unchanged, false);

const secondCheck = engine.checkUnchanged("test.txt", sampleContent);
assert.equal(secondCheck.unchanged, true);
console.log("[3] Unchanged hash detection verified:", secondCheck.hash);

// 4. Range read
const range = engine.readByRange(sampleContent, 2, 2);
assert.equal(range.content, "Línea 2: Parámetros del sistema");
assert.equal(range.lineCount, 1);

// 5. Anchor patch (Surgical edit)
const patchRes = engine.applyAnchorPatch(sampleContent, {
  target: "Parámetros del sistema",
  replacement: "Parámetros actualizados V26"
});
assert.match(patchRes.modified, /Parámetros actualizados V26/);
console.log("[4] Anchor patch verified. Hash before:", patchRes.beforeHash, "after:", patchRes.afterHash);

// 6. Range patch
const rangePatchRes = engine.applyLineRangePatch(sampleContent, {
  startLine: 1,
  endLine: 1,
  replacement: "Línea 1: Nueva cabecera"
});
assert.match(rangePatchRes.modified, /^Línea 1: Nueva cabecera/);

// 7. Lossless Round-Trip: Unicode, Accents, Emoji, Special Chars
const unicodeData = "¡Hola Mundo! 🚀 — Configuración: Ñandú, áéíóú, 中文测试, 100% garantía.";
const stashUnicode = engine.stashSet(unicodeData);
assert.equal(engine.stashGet(stashUnicode.stashId), unicodeData);
const unicodeHash = engine.computeHash(unicodeData);
const retrievedHash = engine.computeHash(engine.stashGet(stashUnicode.stashId));
assert.equal(unicodeHash, retrievedHash, "Lossless Unicode check failed!");
console.log("[5] Lossless Unicode hash verification passed!");

// 8. Large Payload (>1 MB) lossless test
const large1MB = "X".repeat(1024 * 1024 + 500); // > 1MB
const largeStash = engine.stashSet(large1MB);
assert.equal(engine.stashGet(largeStash.stashId).length, large1MB.length);
assert.equal(engine.computeHash(engine.stashGet(largeStash.stashId)), engine.computeHash(large1MB));
console.log("[6] 1MB+ Large string lossless verification passed!");

// 9. Edge cases: empty string, nonexistent anchor error
assert.throws(() => {
  engine.applyAnchorPatch(sampleContent, { target: "NO_EXISTE", replacement: "foo" });
}, /Target content no encontrado/);

console.log("\n[PASS] All compact engine unit tests passed successfully!");
