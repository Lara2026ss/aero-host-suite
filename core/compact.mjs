/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
 * core/compact.mjs — High-Efficiency Token Optimization Engine
 * 
 * ZERO GIMMICKS: NO binary, NO morse, NO hex, NO gzip+base64.
 * Real savings via references, stashes, diffs, handles, range reads, and smart truncation.
 */
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

export class CompactEngine {
  constructor(options = {}) {
    this.mode = options.mode || "auto"; // "off" | "auto" | "aggressive"
    this.stash = new Map(); // stashId -> string content
    this.handles = new Map(); // handleId -> { type, path, meta, createdAt }
    this.fileHashes = new Map(); // normalizedPath -> { hash, readAt }
    this.handleCounter = 0;
    this.maxStashEntries = options.maxStashEntries || 100;
  }

  setMode(mode) {
    if (["off", "auto", "aggressive"].includes(mode)) {
      this.mode = mode;
      return true;
    }
    return false;
  }

  getMode() {
    return this.mode;
  }

  // ── 1. Content Stash ──────────────────────────────────────────────────────────
  
  stashSet(content, customId = null) {
    const text = String(content ?? "");
    const hash = crypto.createHash("sha256").update(text, "utf8").digest("hex").slice(0, 8);
    const stashId = customId || `stash#${hash}`;
    
    if (this.stash.size >= this.maxStashEntries) {
      const oldestKey = this.stash.keys().next().value;
      this.stash.delete(oldestKey);
    }

    this.stash.set(stashId, {
      content: text,
      hash,
      length: text.length,
      createdAt: Date.now()
    });
    return { stashId, length: text.length, hash };
  }

  stashGet(stashId) {
    if (!stashId) return null;
    const entry = this.stash.get(String(stashId).trim());
    return entry ? entry.content : null;
  }

  stashHas(stashId) {
    return this.stash.has(String(stashId).trim());
  }

  resolveContent(input) {
    if (typeof input !== "string") return input;
    const trimmed = input.trim();
    if (trimmed.startsWith("stash#") && this.stash.has(trimmed)) {
      return this.stash.get(trimmed).content;
    }
    return input;
  }

  // ── 2. Handles ───────────────────────────────────────────────────────────────

  createHandle(type, targetPath, meta = {}) {
    this.handleCounter++;
    const prefix = type === "doc" ? "doc" : "f";
    const handleId = `${prefix}#${this.handleCounter}`;
    this.handles.set(handleId, {
      id: handleId,
      type,
      path: path.resolve(targetPath),
      meta,
      createdAt: Date.now()
    });
    return handleId;
  }

  resolveHandle(handleId) {
    if (!handleId || typeof handleId !== "string") return null;
    const entry = this.handles.get(handleId.trim());
    return entry || null;
  }

  releaseHandle(handleId) {
    return this.handles.delete(String(handleId).trim());
  }

  // ── 3. Read Optimization & Hash Cache ────────────────────────────────────────

  computeHash(content) {
    return crypto.createHash("sha256").update(content, "utf8").digest("hex");
  }

  recordRead(filePath, content) {
    const norm = path.resolve(filePath);
    const hash = this.computeHash(content);
    const prev = this.fileHashes.get(norm);
    this.fileHashes.set(norm, { hash, readAt: Date.now() });
    return {
      hash,
      unchanged: prev ? prev.hash === hash : false
    };
  }

  checkUnchanged(filePath, currentContent) {
    if (this.mode === "off") return { unchanged: false };
    const norm = path.resolve(filePath);
    const prev = this.fileHashes.get(norm);
    const currentHash = this.computeHash(currentContent);
    if (prev && prev.hash === currentHash) {
      return {
        unchanged: true,
        hash: currentHash,
        message: "Sin cambios desde la última lectura."
      };
    }
    this.fileHashes.set(norm, { hash: currentHash, readAt: Date.now() });
    return { unchanged: false, hash: currentHash };
  }

  readByRange(content, startLine = 1, endLine = null) {
    const lines = content.split(/\r?\n/);
    const totalLines = lines.length;
    const start = Math.max(1, parseInt(startLine, 10) || 1);
    const end = endLine ? Math.min(totalLines, parseInt(endLine, 10)) : totalLines;
    
    if (start > totalLines) {
      return { lines: [], start, end: totalLines, totalLines, content: "" };
    }
    
    const slice = lines.slice(start - 1, end);
    return {
      start,
      end,
      totalLines,
      lineCount: slice.length,
      content: slice.join("\n")
    };
  }

  getOutline(content, ext = "") {
    const lines = content.split(/\r?\n/);
    const outline = [];
    const lowerExt = ext.toLowerCase();

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmed = line.trim();
      
      // Markdown headings
      if (lowerExt === ".md" && /^#{1,6}\s+/.test(trimmed)) {
        outline.push({ line: i + 1, type: "heading", text: trimmed });
      }
      // JS / TS functions and classes
      else if ([".js", ".mjs", ".ts", ".jsx", ".tsx"].includes(lowerExt)) {
        if (/^(?:export\s+)?(?:async\s+)?function\s+([A-Za-z0-9_$]+)/.test(trimmed) ||
            /^(?:export\s+)?class\s+([A-Za-z0-9_$]+)/.test(trimmed) ||
            /^(?:export\s+)?const\s+([A-Za-z0-9_$]+)\s*=\s*(?:async\s*)?\([^)]*\)\s*=>/.test(trimmed)) {
          outline.push({ line: i + 1, type: "declaration", text: trimmed.slice(0, 100) });
        }
      }
      // Python def / class
      else if (lowerExt === ".py" && /^(?:async\s+)?def\s+([A-Za-z0-9_]+)|class\s+([A-Za-z0-9_]+)/.test(trimmed)) {
        outline.push({ line: i + 1, type: "definition", text: trimmed });
      }
    }
    return { totalLines: lines.length, entriesCount: outline.length, outline };
  }

  // ── 4. Incremental Writes & Patching ──────────────────────────────────────────

  applyAnchorPatch(originalContent, { target, replacement, allowMultiple = false }) {
    if (typeof originalContent !== "string") throw new Error("Contenido original inválido");
    if (!target) throw new Error("El objetivo a reemplazar no puede estar vacío");

    const occurrences = originalContent.split(target).length - 1;
    if (occurrences === 0) {
      throw new Error(`Target content no encontrado en el archivo.`);
    }
    if (occurrences > 1 && !allowMultiple) {
      throw new Error(`Target content encontrado ${occurrences} veces. Use allowMultiple=true o especifique más contexto.`);
    }

    let modified;
    if (allowMultiple) {
      modified = originalContent.replaceAll(target, replacement);
    } else {
      modified = originalContent.replace(target, replacement);
    }

    return {
      modified,
      replacements: allowMultiple ? occurrences : 1,
      beforeHash: this.computeHash(originalContent),
      afterHash: this.computeHash(modified)
    };
  }

  applyLineRangePatch(originalContent, { startLine, endLine, replacement }) {
    const lines = originalContent.split(/\r?\n/);
    const start = parseInt(startLine, 10);
    const end = parseInt(endLine, 10);

    if (isNaN(start) || isNaN(end) || start < 1 || end < start || start > lines.length) {
      throw new Error(`Rango de líneas inválido: [${startLine}, ${endLine}]. Total líneas: ${lines.length}`);
    }

    const replacementLines = replacement ? replacement.split(/\r?\n/) : [];
    const before = lines.slice(0, start - 1);
    const after = lines.slice(end);
    const modifiedLines = [...before, ...replacementLines, ...after];
    const modified = modifiedLines.join("\n");

    return {
      modified,
      replacedLines: end - start + 1,
      newLinesCount: replacementLines.length,
      beforeHash: this.computeHash(originalContent),
      afterHash: this.computeHash(modified)
    };
  }

  insertRelative(originalContent, { anchor, content, position = "after" }) {
    const idx = originalContent.indexOf(anchor);
    if (idx === -1) {
      throw new Error(`Anchor text no encontrado para inserción`);
    }

    let modified;
    if (position === "after") {
      const insertionPoint = idx + anchor.length;
      modified = originalContent.slice(0, insertionPoint) + "\n" + content + originalContent.slice(insertionPoint);
    } else {
      modified = originalContent.slice(0, idx) + content + "\n" + originalContent.slice(idx);
    }

    return {
      modified,
      beforeHash: this.computeHash(originalContent),
      afterHash: this.computeHash(modified)
    };
  }

  // ── 5. Response Truncation & Formatting ───────────────────────────────────────

  truncateOutput(text, maxLines = 80, maxChars = 6000) {
    if (!text) return "";
    const str = String(text);
    if (str.length <= maxChars) {
      const lines = str.split(/\r?\n/);
      if (lines.length <= maxLines) return str;
    }

    const lines = str.split(/\r?\n/);
    const headLines = Math.floor(maxLines * 0.4);
    const tailLines = Math.floor(maxLines * 0.4);
    const omitted = lines.length - (headLines + tailLines);

    if (omitted <= 5) return str;

    const head = lines.slice(0, headLines).join("\n");
    const tail = lines.slice(lines.length - tailLines).join("\n");
    return `${head}\n\n... [${omitted} líneas omitidas por Compact Mode. Use read_range o detail=true para ver más] ...\n\n${tail}`;
  }

  formatResponse({ ok, summary, data, detail = false }) {
    if (detail || this.mode === "off") {
      return { ok, summary, data };
    }
    // Compact mode: Minimal response by default
    return { ok, summary, ...(data?.stashId ? { stashId: data.stashId } : {}), ...(data?.handle ? { handle: data.handle } : {}) };
  }

  estimateTokens(text) {
    if (!text) return 0;
    const str = String(text);
    const wordsAndPunct = str.match(/[\p{L}\p{N}]+|[^\s\p{L}\p{N}]+|\s+/gu);
    return wordsAndPunct ? Math.round(wordsAndPunct.length * 0.9) : Math.ceil(str.length / 3.5);
  }
}

export const compact = new CompactEngine();
