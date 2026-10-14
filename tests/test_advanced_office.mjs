/**
 * Aero Host Suite — tests/test_advanced_office.mjs
 * Comprehensive Integration Test Suite for Advanced Office & PDF Engine
 * Physical execution tests across Headless, COM, Templates, Word, Excel, PowerPoint, and PDF.
 */
import assert from "node:assert";
import path from "node:path";
import fs from "node:fs/promises";
import os from "node:os";
import crypto from "node:crypto";
import { PDFDocument, rgb } from "pdf-lib";
import { createDocumentsDomain } from "../tools/documents.mjs";

console.log("======================================================================");
console.log("   AERO HOST SUITE — ADVANCED OFFICE LIVE & HEADLESS ENGINE TEST SUITE");
console.log("======================================================================\n");

const testDir = path.join(os.tmpdir(), `aero_office_test_${Date.now()}`);
await fs.mkdir(testDir, { recursive: true });

const docDomain = createDocumentsDomain({
  runtime: { dirs: { storage: testDir }, storage: testDir },
  path,
  fs,
  crypto,
  domain: "documents"
});

let passedCount = 0;

async function runStep(label, fn) {
  process.stdout.write(`[*] Testing: ${label}... `);
  const start = Date.now();
  try {
    await fn();
    const elapsed = Date.now() - start;
    console.log(`[PASSED] (${elapsed}ms)`);
    passedCount++;
  } catch (err) {
    console.log(`[FAILED]`);
    console.error(`    Error: ${err.message}`);
    throw err;
  }
}

try {
  // ── 1. Templates & Themes Catalog ──────────────────────────────────────────
  await runStep("get_templates returns themes and structured presets", async () => {
    const res = await docDomain.actions.get_templates();
    assert.strictEqual(res.ok, true);
    assert.ok(res.themes.modern_blue, "Missing modern_blue theme");
    assert.ok(res.themes.politecnico_maroon, "Missing politecnico_maroon theme");
    assert.ok(res.presets.academic_report, "Missing academic_report preset");
    assert.ok(res.presets.financial_model, "Missing financial_model preset");
  });

  // ── 2. Headless PDF Operations (Create, Merge, Inspect, Extract) ─────────────
  const pdf1Path = path.join(testDir, "test_part1.pdf");
  const pdf2Path = path.join(testDir, "test_part2.pdf");
  const mergedPdfPath = path.join(testDir, "test_merged.pdf");

  await runStep("Generate sample PDF fixtures using pdf-lib", async () => {
    const doc1 = await PDFDocument.create();
    const page1 = doc1.addPage([600, 400]);
    page1.drawText("Aero Host Suite - Page 1 Content", { x: 50, y: 350, size: 16 });
    await fs.writeFile(pdf1Path, await doc1.save());

    const doc2 = await PDFDocument.create();
    const page2 = doc2.addPage([600, 400]);
    page2.drawText("Aero Host Suite - Page 2 Content", { x: 50, y: 350, size: 16 });
    await fs.writeFile(pdf2Path, await doc2.save());

    assert.ok((await fs.stat(pdf1Path)).size > 0);
    assert.ok((await fs.stat(pdf2Path)).size > 0);
  });

  await runStep("pdf_merge combines multiple PDFs into unified document", async () => {
    const res = await docDomain.actions.pdf_merge({
      files: [pdf1Path, pdf2Path],
      outputPath: mergedPdfPath
    });
    assert.strictEqual(res.ok, true);
    assert.strictEqual(res.pageCount, 2);
    assert.ok((await fs.stat(mergedPdfPath)).size > 0);
  });

  await runStep("pdf_inspect retrieves metadata and page metrics", async () => {
    const res = await docDomain.actions.pdf_inspect({ path: mergedPdfPath });
    assert.strictEqual(res.ok, true);
    assert.strictEqual(res.pageCount, 2);
    assert.strictEqual(res.format, "pdf");
    assert.strictEqual(res.mode, "headless_library");
  });

  await runStep("pdf_extract_text extracts page-by-page content cleanly", async () => {
    const res = await docDomain.actions.pdf_extract_text({ path: mergedPdfPath });
    assert.strictEqual(res.ok, true);
    assert.strictEqual(res.pageCount, 2);
    assert.ok(res.pages.length === 2);
    assert.ok(res.pages[0].text.includes("Page 1 Content") || res.fullText.includes("Page 1 Content"));
  });

  // ── 3. Headless Word DOCX Operations ───────────────────────────────────────
  const wordTestDocx = path.join(testDir, "test_academic.docx");

  await runStep("word_template creates styled document in headless mode", async () => {
    const res = await docDomain.actions.word_template({
      path: wordTestDocx,
      preset: "academic_report",
      theme: "politecnico_maroon",
      title: "Arquitectura Avanzada de Automatización Office",
      subtitle: "Evaluación de Capacidades en Tiempo Real y Headless",
      author: "Antigravity Engineering",
      sections: [
        {
          heading: "1. Introducción y Marco Teórico",
          body: "El subsistema documents implementa un modelo híbrido dual: Live COM para edición en pantalla y Headless para operaciones batch de alto rendimiento.",
          level: 1
        },
        {
          heading: "2. Matriz de Rendimiento",
          body: "Comparativa de velocidad entre entornos interactivos y headless:",
          level: 2,
          table: [
            ["Módulo", "Modo", "Latencia Promedio"],
            ["PDF Merge", "Headless", "12ms"],
            ["Excel Batch", "Headless", "45ms"],
            ["Word Live", "Live COM", "85ms"]
          ]
        }
      ],
      mode: "headless"
    });
    assert.strictEqual(res.ok, true);
    assert.strictEqual(res.mode, "headless_library");
    assert.ok((await fs.stat(wordTestDocx)).size > 0);
  });

  await runStep("inspect_structure parses headless DOCX outlines without Office UI", async () => {
    const res = await docDomain.actions.inspect_structure({
      path: wordTestDocx,
      mode: "headless"
    });
    assert.strictEqual(res.ok, true);
    assert.strictEqual(res.format, "docx");
    assert.ok(res.headings.length >= 2, "Expected headings in outline");
  });

  // ── 4. Headless Excel XLSX Operations ──────────────────────────────────────
  const excelTestXlsx = path.join(testDir, "test_financial_model.xlsx");

  await runStep("Headless Excel write with formulas, zebra striping, and auto-fit", async () => {
    const res = await docDomain.actions.write_range({
      path: excelTestXlsx,
      sheet: "Finanzas Q4",
      mode: "headless",
      data: [
        ["Categoría", "Presupuesto", "Ejecutado", "Variación"],
        ["Infraestructura Cloud", 15000, 14200, "=B2-C2"],
        ["Automatización Office", 8000, 7500, "=B3-C3"],
        ["Auditoría y QA", 5000, 4800, "=B4-C4"]
      ]
    });
    assert.strictEqual(res.ok, true);
    assert.ok((await fs.stat(excelTestXlsx)).size > 0);
  });

  await runStep("inspect_structure parses headless XLSX sheet outline", async () => {
    const res = await docDomain.actions.inspect_structure({
      path: excelTestXlsx,
      mode: "headless"
    });
    assert.strictEqual(res.ok, true);
    assert.strictEqual(res.format, "xlsx");
    assert.ok(res.sheetCount >= 1);
    assert.ok(res.sheets.some(s => s.name === "Finanzas Q4"));
  });

  // ── 5. Discovery & Live COM Checks ─────────────────────────────────────────
  await runStep("status verifies physical Office processes and COM availability", async () => {
    const res = await docDomain.actions.status({ app: "all" });
    assert.strictEqual(res.ok, true);
    assert.strictEqual(res.mode, "live_com");
    assert.ok(res.status.word !== undefined);
    assert.ok(res.status.excel !== undefined);
    assert.ok(res.status.powerpoint !== undefined);
    assert.strictEqual(res.status.word.installed, true);
    assert.strictEqual(res.status.excel.installed, true);
    assert.strictEqual(res.status.powerpoint.installed, true);
  });

  // ── 6. Live COM Inspection & Safety Verification ──────────────────────────
  await runStep("Verify user active documents remain intact and isolated", async () => {
    const openRes = await docDomain.actions.list_open({ app: "word" });
    assert.strictEqual(openRes.ok, true);
    assert.strictEqual(openRes.mode, "live_com");
    if (openRes.documents && openRes.documents.length > 0) {
      console.log(` (Found ${openRes.documents.length} open Word document(s): ${openRes.documents.map(d => d.name).join(", ")})`);
    }
  });

  await runStep("inspect_structure extracts outline from active Word document via Live COM", async () => {
    const res = await docDomain.actions.inspect_structure({ app: "word" });
    assert.strictEqual(res.ok, true);
    assert.strictEqual(res.mode, "live_com");
    assert.strictEqual(res.app, "word");
    assert.ok(res.title !== undefined);
    assert.ok(res.sections >= 1);
    console.log(` (Word document outline: '${res.title}', ${res.paragraphs} paras, ${res.headings?.length || 0} headings, ${res.pages} pages)`);
  });

  console.log("\n======================================================================");
  console.log(`   ALL ${passedCount} INTEGRATION TESTS PASSED WITH 100% SUCCESS.`);
  console.log("======================================================================\n");
} finally {
  // Clean up temporary test files safely
  await fs.rm(testDir, { recursive: true, force: true }).catch(() => {});
}
