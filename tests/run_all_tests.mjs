/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * tests/run_all_tests.mjs — Master Test Suite Runner
 * Runs all official V26-0.1 test suites with physical execution assertions.
 */
import { exec } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const TEST_SUITES = [
  { name: "1. Version & Semver Ordering", file: "test_version.mjs" },
  { name: "2. Compact Mode & Surgical Diffs", file: "test_compact.mjs" },
  { name: "3. Security Permissions & AppUserModelID", file: "test_permissions.mjs" },
  { name: "4. Multi-Step Shortcuts & Rollback", file: "test_shortcuts.mjs" },
  { name: "5. Feedback Engine & Privacy Sanitization", file: "test_feedback.mjs" },
  { name: "6. Autonomous Remote Updater & Rollback", file: "test_updater_real.mjs" },
  { name: "7. Reconstructed Runtime & All 11 Tools", file: "test_runtime_and_tools.mjs" },
  { name: "8. Physical Token Savings Benchmark", file: "measure_token_savings.mjs" }
];

console.log("======================================================================");
console.log("   FLUXER XZ - PUBLIC RELEASE V26-0.1 :: MASTER TEST SUITE RUNNER");
console.log("======================================================================\n");

let passed = 0;
let failed = 0;
const results = [];

for (const suite of TEST_SUITES) {
  const filePath = path.join(__dirname, suite.file);
  process.stdout.write(`[*] Running ${suite.name} (${suite.file})... `);
  const start = Date.now();

  const code = await new Promise((resolve) => {
    exec(`node "${filePath}"`, (err, stdout, stderr) => {
      const elapsed = Date.now() - start;
      if (err) {
        console.log(`[FAILED] (${elapsed}ms)`);
        if (stdout) console.log(stdout.trim());
        if (stderr) console.error(stderr.trim());
        resolve({ code: err.code || 1, elapsed, error: stderr || err.message });
      } else {
        console.log(`[PASSED] (${elapsed}ms)`);
        resolve({ code: 0, elapsed });
      }
    });
  });

  if (code.code === 0) {
    passed++;
    results.push({ name: suite.name, status: "PASSED", elapsed: code.elapsed });
  } else {
    failed++;
    results.push({ name: suite.name, status: "FAILED", elapsed: code.elapsed, error: code.error });
  }
}

console.log("\n======================================================================");
console.log(`   TEST SUMMARY: ${passed} PASSED / ${failed} FAILED (${TEST_SUITES.length} TOTAL)`);
console.log("======================================================================");

for (const r of results) {
  const mark = r.status === "PASSED" ? "[OK]" : "[FAIL]";
  console.log(`  ${mark} ${r.name.padEnd(46)} ${r.elapsed}ms`);
}

if (failed > 0) {
  console.error(`\n[!] Test run completed with ${failed} failure(s).`);
  process.exit(1);
} else {
  console.log("\n[SUCCESS] ALL TEST SUITES PASSED PHYSICALLY WITH EXIT CODE 0.");
  process.exit(0);
}
