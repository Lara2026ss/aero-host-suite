import assert from "node:assert/strict";
import { getVersion, getSemver, compareVersions, FULL_VERSION_STRING } from "../core/version.mjs";

console.log("=== TESTING VERSION ENGINE ===");
console.log("Current version:", getVersion());
console.log("Semver mapped:", getSemver());
console.log("Full version string:", FULL_VERSION_STRING);

assert.equal(getVersion(), "V26-0.1");
assert.equal(getSemver(), "26.0.1");

// Ordering test specified in prompt:
// V26-0.1 < V26-0.2 < V26-1.0 < V27-0.1
assert.equal(compareVersions("V26-0.1", "V26-0.2"), -1, "V26-0.1 should be < V26-0.2");
assert.equal(compareVersions("V26-0.2", "V26-1.0"), -1, "V26-0.2 should be < V26-1.0");
assert.equal(compareVersions("V26-1.0", "V27-0.1"), -1, "V26-1.0 should be < V27-0.1");

// Transitivity check
assert.equal(compareVersions("V26-0.1", "V27-0.1"), -1, "V26-0.1 should be < V27-0.1");
assert.equal(compareVersions("V26-0.1", "V26-0.1"), 0, "Same version should be 0");
assert.equal(compareVersions("V26-0.2", "V26-0.1"), 1, "V26-0.2 should be > V26-0.1");

console.log("[PASS] Version ordering test: V26-0.1 < V26-0.2 < V26-1.0 < V27-0.1 verified!");
