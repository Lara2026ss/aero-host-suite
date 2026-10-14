/**
 * Aero Host Suite — tools/documents/headless.mjs
 * High-performance Headless Document Engine (DOCX, XLSX, PDF)
 * Operates without Microsoft Office UI/process dependencies.
 */
import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import * as docx from "docx";
import { PDFDocument } from "pdf-lib";
import mammoth from "mammoth";
import { extractText } from "unpdf";

export class HeadlessDocumentEngine {
  /**
   * Inspect PDF metadata and page count.
   */
  async pdfInspect(filePath) {
    const absPath = path.resolve(filePath);
    const buffer = await fs.readFile(absPath);
    const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
    
    return {
      ok: true,
      mode: "headless_library",
      format: "pdf",
      path: absPath,
      pageCount: pdfDoc.getPageCount(),
      title: pdfDoc.getTitle() || "",
      author: pdfDoc.getAuthor() || "",
      subject: pdfDoc.getSubject() || "",
      creator: pdfDoc.getCreator() || "",
      producer: pdfDoc.getProducer() || "",
      creationDate: pdfDoc.getCreationDate() ? pdfDoc.getCreationDate().toISOString() : null,
      modificationDate: pdfDoc.getModificationDate() ? pdfDoc.getModificationDate().toISOString() : null,
      sizeBytes: buffer.length,
    };
  }

  /**
   * Extract text and page layout from PDF.
   */
  async pdfExtractText(filePath, options = {}) {
    const absPath = path.resolve(filePath);
    const buffer = await fs.readFile(absPath);
    const uint8 = new Uint8Array(buffer);
    
    const maxPages = options.maxPages ? Number(options.maxPages) : 30;
    const startPage = options.startPage ? Math.max(1, Number(options.startPage)) : 1;
    
    const result = await extractText(uint8, { mergePages: false });
    const totalPages = result.totalPages || result.text.length;
    
    const endPage = options.endPage ? Math.min(totalPages, Number(options.endPage)) : Math.min(totalPages, startPage + maxPages - 1);
    
    const pageTexts = [];
    for (let p = startPage; p <= endPage; p++) {
      const pageIndex = p - 1;
      const text = (result.text && result.text[pageIndex]) ? result.text[pageIndex].trim() : "";
      pageTexts.push({
        pageNumber: p,
        text,
        charCount: text.length,
      });
    }

    return {
      ok: true,
      mode: "headless_library",
      format: "pdf",
      path: absPath,
      totalPages,
      pageCount: totalPages,
      startPage,
      endPage,
      extractedCount: pageTexts.length,
      pages: pageTexts,
      fullText: pageTexts.map((p) => `--- [Page ${p.pageNumber}] ---\n${p.text}`).join("\n\n"),
    };
  }

  /**
   * Merge multiple PDF files into a single unified document.
   */
  async pdfMerge(filePaths, outputPath) {
    if (!Array.isArray(filePaths) || filePaths.length < 2) {
      throw new Error("pdfMerge requires an array of at least 2 file paths");
    }
    const absOutputPath = path.resolve(outputPath);
    await fs.mkdir(path.dirname(absOutputPath), { recursive: true });

    const mergedPdf = await PDFDocument.create();
    const sourceStats = [];

    for (const src of filePaths) {
      const absSrc = path.resolve(src);
      const buffer = await fs.readFile(absSrc);
      const doc = await PDFDocument.load(buffer, { ignoreEncryption: true });
      const indices = doc.getPageIndices();
      const copiedPages = await mergedPdf.copyPages(doc, indices);
      for (const page of copiedPages) {
        mergedPdf.addPage(page);
      }
      sourceStats.push({ path: absSrc, pageCount: indices.length });
    }

    const mergedBytes = await mergedPdf.save();
    await fs.writeFile(absOutputPath, mergedBytes);

    return {
      ok: true,
      mode: "headless_library",
      format: "pdf",
      outputPath: absOutputPath,
      totalPages: mergedPdf.getPageCount(),
      pageCount: mergedPdf.getPageCount(),
      sourceFiles: sourceStats,
      sizeBytes: mergedBytes.length,
    };
  }

  /**
   * Read Excel workbook sheets, cell data, dimensions, and formulas without Excel.exe.
   */
  async excelRead(filePath, options = {}) {
    const absPath = path.resolve(filePath);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(absPath);

    const sheetsSummary = wb.worksheets.map((ws) => ({
      id: ws.id,
      name: ws.name,
      rowCount: ws.rowCount,
      columnCount: ws.columnCount,
      actualRowCount: ws.actualRowCount,
      actualColumnCount: ws.actualColumnCount,
    }));

    const targetSheet = options.sheet
      ? wb.getWorksheet(options.sheet)
      : wb.worksheets[0];

    if (!targetSheet) {
      return {
        ok: true,
        mode: "headless_library",
        format: "xlsx",
        path: absPath,
        sheets: sheetsSummary,
        values: [],
        formulas: [],
        message: "No worksheet matched target",
      };
    }

    const values = [];
    const formulas = [];
    const maxRows = options.maxRows ? Number(options.maxRows) : 200;

    targetSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > maxRows) return;
      const rowVals = [];
      const rowForms = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        rowVals.push(cell.value);
        rowForms.push(cell.formula || null);
      });
      values.push(rowVals);
      formulas.push(rowForms);
    });

    return {
      ok: true,
      mode: "headless_library",
      format: "xlsx",
      path: absPath,
      sheet: targetSheet.name,
      sheetCount: sheetsSummary.length,
      sheets: sheetsSummary,
      rowsRead: values.length,
      values,
      formulas,
    };
  }

  /**
   * Create or update Excel workbook with rich styles, formulas, and auto-sizing.
   */
  async excelWrite(filePath, options = {}) {
    const absPath = path.resolve(filePath);
    await fs.mkdir(path.dirname(absPath), { recursive: true });

    const wb = new ExcelJS.Workbook();
    let fileExists = false;
    try {
      await fs.access(absPath);
      fileExists = true;
    } catch {}

    if (fileExists && !options.overwrite) {
      try {
        await wb.xlsx.readFile(absPath);
      } catch {
        // Fallback to fresh workbook if corrupt or unreadable
      }
    }

    const sheetName = options.sheet || options.sheetName || "Sheet1";
    let ws = wb.getWorksheet(sheetName);
    if (!ws) {
      ws = wb.addWorksheet(sheetName);
    }

    const startCellStr = options.startCell || "A1";
    const startCell = ws.getCell(startCellStr);
    const startRow = startCell.row;
    const startCol = startCell.col;

    // Optional Headers
    if (Array.isArray(options.headers) && options.headers.length > 0) {
      const headerRow = ws.getRow(startRow);
      options.headers.forEach((h, i) => {
        const cell = headerRow.getCell(startCol + i);
        cell.value = h;
        cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11, name: "Calibri" };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: options.headerBgColor ? options.headerBgColor.replace("#", "") : "FF1F4E79" },
        };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
          top: { style: "thin" },
          bottom: { style: "medium" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
      });
      headerRow.height = 24;
    }

    // Write Data Matrix
    const dataOffset = Array.isArray(options.headers) && options.headers.length > 0 ? 1 : 0;
    const data = Array.isArray(options.data) ? options.data : [];

    data.forEach((rowVals, rIdx) => {
      const currentRowNum = startRow + dataOffset + rIdx;
      const row = ws.getRow(currentRowNum);
      if (Array.isArray(rowVals)) {
        rowVals.forEach((val, cIdx) => {
          const cell = row.getCell(startCol + cIdx);
          if (typeof val === "string" && val.startsWith("=")) {
            cell.value = { formula: val.slice(1) };
          } else {
            cell.value = val;
          }

          // Optional zebra striping
          if (options.zebra && rIdx % 2 === 1) {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: options.zebraColor ? options.zebraColor.replace("#", "") : "FFF2F2F2" },
            };
          }

          // Border
          if (options.borders) {
            cell.border = {
              top: { style: "thin", color: { argb: "FFD9D9D9" } },
              bottom: { style: "thin", color: { argb: "FFD9D9D9" } },
              left: { style: "thin", color: { argb: "FFD9D9D9" } },
              right: { style: "thin", color: { argb: "FFD9D9D9" } },
            };
          }
        });
      }
    });

    // Auto-fit Column Widths
    if (options.autoFit !== false) {
      ws.columns.forEach((column) => {
        let maxLen = 10;
        column.eachCell({ includeEmpty: false }, (cell) => {
          const len = cell.value ? String(cell.value).length : 0;
          if (len > maxLen) maxLen = Math.min(len + 3, 50);
        });
        column.width = maxLen;
      });
    }

    await wb.xlsx.writeFile(absPath);

    return {
      ok: true,
      mode: "headless_library",
      format: "xlsx",
      path: absPath,
      sheet: ws.name,
      totalSheets: wb.worksheets.length,
      writtenRows: data.length + dataOffset,
      sizeBytes: (await fs.stat(absPath)).size,
    };
  }

  /**
   * Extract text, heading structure, and word count from DOCX using Mammoth.
   */
  async docxExtract(filePath) {
    const absPath = path.resolve(filePath);
    const buffer = await fs.readFile(absPath);
    
    const rawResult = await mammoth.extractRawText({ buffer });
    const text = rawResult.value.trim();

    // Structural outline detection from lines
    const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    const headings = [];
    lines.forEach((line, idx) => {
      // Heuristic detection: short uppercase or numbered lines
      if (line.length < 80 && (/^(?:[0-9]+(?:\.[0-9]+)*|CAP[ÍI]TULO|SECCI[ÓO]N|INTRODUCCI|CONCLUSI)/i.test(line) || /^[A-ZÁÉÍÓÚ\s]{4,}$/.test(line))) {
        headings.push({ index: idx + 1, text: line });
      }
    });

    const wordCount = text ? text.split(/\s+/).length : 0;

    return {
      ok: true,
      mode: "headless_library",
      format: "docx",
      path: absPath,
      charCount: text.length,
      wordCount,
      paragraphCount: lines.length,
      headings,
      preview: text.slice(0, 2000),
    };
  }

  /**
   * Create a structured professional DOCX file headlessly with docx package.
   */
  async docxCreate(filePath, options = {}) {
    const absPath = path.resolve(filePath);
    await fs.mkdir(path.dirname(absPath), { recursive: true });

    const docChildren = [];

    // Title
    if (options.title) {
      docChildren.push(
        new docx.Paragraph({
          text: options.title,
          heading: docx.HeadingLevel.TITLE,
          spacing: { after: 240 },
          alignment: docx.AlignmentType.CENTER,
        })
      );
    }

    // Subtitle
    if (options.subtitle) {
      docChildren.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: options.subtitle,
              italics: true,
              size: 24,
              color: "666666",
            }),
          ],
          spacing: { after: 360 },
          alignment: docx.AlignmentType.CENTER,
        })
      );
    }

    // Sections
    const sections = Array.isArray(options.sections) ? options.sections : [];
    sections.forEach((sec) => {
      if (sec.heading) {
        docChildren.push(
          new docx.Paragraph({
            text: sec.heading,
            heading: sec.level === 2 ? docx.HeadingLevel.HEADING_2 : docx.HeadingLevel.HEADING_1,
            spacing: { before: 240, after: 120 },
          })
        );
      }

      if (sec.content || sec.body) {
        const bodyText = sec.content || sec.body;
        docChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: bodyText, size: 22 })],
            spacing: { after: 180 },
            alignment: docx.AlignmentType.JUSTIFIED,
          })
        );
      }

      // Tables inside section
      if (Array.isArray(sec.table) && sec.table.length > 0) {
        const tableRows = sec.table.map((rowArr, rIdx) => {
          const cells = rowArr.map((cellText) => {
            return new docx.TableCell({
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: String(cellText),
                      bold: rIdx === 0,
                      color: rIdx === 0 ? "FFFFFF" : "000000",
                    }),
                  ],
                }),
              ],
              shading: rIdx === 0 ? { fill: options.themeColor || "1F4E79" } : undefined,
            });
          });
          return new docx.TableRow({ children: cells });
        });

        docChildren.push(new docx.Table({ rows: tableRows }));
      }
    });

    const doc = new docx.Document({
      sections: [{ children: docChildren }],
    });

    const buffer = await docx.Packer.toBuffer(doc);
    await fs.writeFile(absPath, buffer);

    return {
      ok: true,
      mode: "headless_library",
      format: "docx",
      path: absPath,
      sizeBytes: buffer.length,
      sectionsCount: sections.length,
    };
  }
}

export const headlessEngine = new HeadlessDocumentEngine();
