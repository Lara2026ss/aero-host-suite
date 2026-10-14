/**
 * Aero Host Suite — tools/documents.mjs
 * Capability: documents (Microsoft Word, PowerPoint, Excel & PDF Automation)
 * Provides real-time LIVE COM and HEADLESS document manipulation for Windows 11.
 *
 * @version 1.0.0
 */
import os from "node:os";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { headlessEngine } from "./documents/headless.mjs";
import { THEMES, TEMPLATE_PRESETS } from "./documents/templates.mjs";

const execFileAsync = promisify(execFile);

export function createDocumentsDomain({ runtime, path, fs, crypto, domain }) {
  const currentDir = path.dirname(fileURLToPath(import.meta.url));
  const scriptPath = path.join(currentDir, "documents", "office-com.ps1");

  async function invokeOfficeCom(action, params = {}) {
    const tempBase = runtime.dirs?.storage || runtime.storage || os.tmpdir();
    const cacheDir = path.join(tempBase, "cache");
    await fs.mkdir(cacheDir, { recursive: true }).catch(() => {});
    
    const payloadFile = path.join(cacheDir, `office_ipc_${crypto.randomUUID()}.json`);
    const payload = { action, params };

    try {
      await fs.writeFile(payloadFile, JSON.stringify(payload), "utf8");
      const { stdout, stderr } = await execFileAsync(
        "powershell.exe",
        [
          "-NoProfile",
          "-NonInteractive",
          "-ExecutionPolicy",
          "Bypass",
          "-File",
          scriptPath,
          "-PayloadFile",
          payloadFile,
        ],
        { timeout: 60000, windowsHide: true }
      );

      const raw = (stdout || "").trim();
      if (!raw) {
        if (stderr) {
          return { ok: false, error: `PowerShell COM failure: ${stderr.trim()}` };
        }
        return { ok: false, error: "Empty response from Office COM runner" };
      }

      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === "object" && parsed.ok && !parsed.mode) {
          parsed.mode = "live_com";
        }
        return parsed;
      } catch (e) {
        return { ok: false, error: `Malformed JSON from Office COM: ${raw.slice(0, 200)}`, details: e.message };
      }
    } catch (err) {
      return {
        ok: false,
        error: err.message,
        code: err.code || "COM_EXECUTION_ERROR",
        timedOut: err.killed || false,
      };
    } finally {
      await fs.unlink(payloadFile).catch(() => {});
    }
  }

  const actions = {
    // ── 1. Discovery Actions ───────────────────────────────────────────────────
    status: async (args = {}) => {
      const app = args.app || args.targetApp || "all";
      return invokeOfficeCom("status", { app });
    },

    list_open: async (args = {}) => {
      const app = args.app || "all";
      return invokeOfficeCom("list_open", { app });
    },

    get_active: async (args = {}) => {
      const app = args.app || "word";
      return invokeOfficeCom("get_active", { app });
    },

    inspect_structure: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const app = args.app || "";
      if (targetPath && targetPath.toLowerCase().endsWith(".pdf")) {
        return headlessEngine.pdfInspect(targetPath);
      }
      if (args.mode === "headless" && targetPath && targetPath.toLowerCase().endsWith(".xlsx")) {
        const readRes = await headlessEngine.excelRead(targetPath, { maxRows: 10 });
        if (!readRes.ok) return readRes;
        return {
          ok: true,
          mode: "headless_library",
          format: "xlsx",
          path: targetPath,
          sheetCount: readRes.sheetCount || readRes.sheets?.length || 0,
          sheets: (readRes.sheets || []).map(s => ({ name: s.name, rowCount: s.rowCount, columnCount: s.columnCount })),
        };
      }
      if (args.mode === "headless" && targetPath && targetPath.toLowerCase().endsWith(".docx")) {
        const extractRes = await headlessEngine.docxExtract(targetPath);
        if (!extractRes.ok) return extractRes;
        return {
          ok: true,
          mode: "headless_library",
          format: "docx",
          path: targetPath,
          headings: extractRes.headings,
          summary: extractRes.summary,
        };
      }
      return invokeOfficeCom("inspect_structure", { path: targetPath, app });
    },

    // ── 2. Lifecycle Actions ───────────────────────────────────────────────────
    create: async (args = {}) => {
      const app = args.app || "word";
      const targetPath = args.path || args.targetPath || args.file || "";
      const template = args.template || "";
      const visible = args.visible !== undefined ? Boolean(args.visible) : true;
      const overwrite = args.overwrite !== undefined ? Boolean(args.overwrite) : false;
      const forceNew = args.forceNew !== undefined ? Boolean(args.forceNew) : false;

      return invokeOfficeCom("create", { app, path: targetPath, template, visible, overwrite, forceNew });
    },

    open: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const visible = args.visible !== undefined ? Boolean(args.visible) : true;
      const readOnly = args.readOnly !== undefined ? Boolean(args.readOnly) : false;

      if (!targetPath) {
        return { ok: false, error: "Parameter 'path' is required to open a document" };
      }

      return invokeOfficeCom("open", { path: targetPath, visible, readOnly });
    },

    save: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const app = args.app || "";
      return invokeOfficeCom("save", { path: targetPath, app });
    },

    save_as: async (args = {}) => {
      const targetPath = args.targetPath || args.destination || "";
      const sourcePath = args.path || args.sourcePath || "";

      if (!targetPath) {
        return { ok: false, error: "Parameter 'targetPath' is required for save_as" };
      }

      return invokeOfficeCom("save_as", { path: sourcePath, targetPath });
    },

    close: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const saveChanges = args.saveChanges !== undefined ? Boolean(args.saveChanges) : true;
      return invokeOfficeCom("close", { path: targetPath, saveChanges });
    },

    export_pdf: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const outputPath = args.outputPath || args.pdfPath || "";

      if (!targetPath) {
        return { ok: false, error: "Parameter 'path' is required for export_pdf" };
      }

      return invokeOfficeCom("export_pdf", { path: targetPath, outputPath });
    },

    read: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const maxChars = args.maxChars ? Number(args.maxChars) : 4000;
      if (targetPath && targetPath.toLowerCase().endsWith(".pdf")) {
        return headlessEngine.pdfExtractText(targetPath, { maxPages: args.maxPages || 20 });
      }
      if (args.mode === "headless" && targetPath && targetPath.toLowerCase().endsWith(".docx")) {
        return headlessEngine.docxExtract(targetPath);
      }
      return invokeOfficeCom("read", { path: targetPath, maxChars });
    },

    insert_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const text = args.text !== undefined ? String(args.text) : (args.content !== undefined ? String(args.content) : "");
      const position = args.position || "end";
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const italic = args.italic !== undefined ? Boolean(args.italic) : null;
      const underline = args.underline !== undefined ? Boolean(args.underline) : null;
      const fontSize = args.fontSize || args.size ? Number(args.fontSize || args.size) : null;
      const fontName = args.fontName || args.font || null;
      const color = args.color || null;
      const alignment = args.alignment || args.align || null;
      const lineSpacing = args.lineSpacing ? Number(args.lineSpacing) : null;
      const spaceAfter = args.spaceAfter !== undefined ? Number(args.spaceAfter) : null;
      const spaceBefore = args.spaceBefore !== undefined ? Number(args.spaceBefore) : null;
      const pageBreak = Boolean(args.pageBreak);
      const style = args.style || null;
      const bullet = Boolean(args.bullet);
      const save = Boolean(args.save);

      if (!text && !pageBreak) return { ok: false, error: "Parameter 'text' or 'content' is required" };

      return invokeOfficeCom("insert_text", {
        path: targetPath,
        text,
        position,
        bold,
        italic,
        underline,
        fontSize,
        fontName,
        color,
        alignment,
        lineSpacing,
        spaceAfter,
        spaceBefore,
        pageBreak,
        style,
        bullet,
        save,
      });
    },

    replace_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const search = args.search || args.find || args.target || "";
      const replace = args.replace !== undefined ? String(args.replace) : (args.replacement !== undefined ? String(args.replacement) : (args.content !== undefined ? String(args.content) : ""));
      const matchCase = Boolean(args.matchCase);
      const matchWholeWord = Boolean(args.matchWholeWord);
      const save = Boolean(args.save);

      if (!search) return { ok: false, error: "Parameter 'search' or 'find' is required" };

      return invokeOfficeCom("replace_text", { path: targetPath, search, replace, matchCase, matchWholeWord, save });
    },

    clear_document: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const save = Boolean(args.save);
      return invokeOfficeCom("clear_document", { path: targetPath, save });
    },

    append_section: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const heading = args.heading || args.title || "";
      const body = args.body || args.content || "";
      const headingLevel = args.headingLevel ? Number(args.headingLevel) : 1;
      const headingColor = args.headingColor || args.titleColor || null;
      const headingAlignment = args.headingAlignment || args.titleAlignment || null;
      const bodyColor = args.bodyColor || null;
      const bodyAlignment = args.bodyAlignment || args.alignment || null;
      const fontSize = args.fontSize || args.size ? Number(args.fontSize || args.size) : null;
      const lineSpacing = args.lineSpacing ? Number(args.lineSpacing) : null;
      const pageBreak = Boolean(args.pageBreak);
      const save = Boolean(args.save);

      if (!heading) return { ok: false, error: "Parameter 'heading' or 'title' is required" };

      return invokeOfficeCom("append_section", {
        path: targetPath,
        heading,
        body,
        headingLevel,
        headingColor,
        headingAlignment,
        bodyColor,
        bodyAlignment,
        fontSize,
        lineSpacing,
        pageBreak,
        save,
      });
    },

    format_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const search = args.search || args.text || args.target || "";
      const all = Boolean(args.all);
      const allOccurrences = args.allOccurrences !== undefined ? Boolean(args.allOccurrences) : true;
      const alignment = args.alignment || args.align || null;
      const fontSize = args.fontSize || args.size ? Number(args.fontSize || args.size) : null;
      const fontName = args.fontName || args.font || null;
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const italic = args.italic !== undefined ? Boolean(args.italic) : null;
      const underline = args.underline !== undefined ? Boolean(args.underline) : null;
      const color = args.color || null;
      const lineSpacing = args.lineSpacing ? Number(args.lineSpacing) : null;
      const spaceBefore = args.spaceBefore !== undefined ? Number(args.spaceBefore) : null;
      const spaceAfter = args.spaceAfter !== undefined ? Number(args.spaceAfter) : null;
      const style = args.style || null;
      const save = Boolean(args.save);

      return invokeOfficeCom("format_text", {
        path: targetPath,
        search,
        all,
        allOccurrences,
        alignment,
        fontSize,
        fontName,
        bold,
        italic,
        underline,
        color,
        lineSpacing,
        spaceBefore,
        spaceAfter,
        style,
        save,
      });
    },

    format_paragraphs: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const start = args.start ? Number(args.start) : 1;
      const end = args.end ? Number(args.end) : null;
      const all = Boolean(args.all);
      const alignment = args.alignment || args.align || null;
      const fontSize = args.fontSize || args.size ? Number(args.fontSize || args.size) : null;
      const fontName = args.fontName || args.font || null;
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const italic = args.italic !== undefined ? Boolean(args.italic) : null;
      const underline = args.underline !== undefined ? Boolean(args.underline) : null;
      const color = args.color || null;
      const lineSpacing = args.lineSpacing ? Number(args.lineSpacing) : null;
      const spaceBefore = args.spaceBefore !== undefined ? Number(args.spaceBefore) : null;
      const spaceAfter = args.spaceAfter !== undefined ? Number(args.spaceAfter) : null;
      const style = args.style || null;
      const save = Boolean(args.save);

      return invokeOfficeCom("format_paragraphs", {
        path: targetPath,
        start,
        end,
        all,
        alignment,
        fontSize,
        fontName,
        bold,
        italic,
        underline,
        color,
        lineSpacing,
        spaceBefore,
        spaceAfter,
        style,
        save,
      });
    },

    page_setup: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const paperSize = args.paperSize || args.size || "letter";
      const orientation = args.orientation || "portrait";
      const margins = args.margins || "normal";
      const topMargin = args.topMargin || args.top ? Number(args.topMargin || args.top) : null;
      const bottomMargin = args.bottomMargin || args.bottom ? Number(args.bottomMargin || args.bottom) : null;
      const leftMargin = args.leftMargin || args.left ? Number(args.leftMargin || args.left) : null;
      const rightMargin = args.rightMargin || args.right ? Number(args.rightMargin || args.right) : null;
      const save = Boolean(args.save);

      return invokeOfficeCom("page_setup", {
        path: targetPath,
        paperSize,
        orientation,
        margins,
        topMargin,
        bottomMargin,
        leftMargin,
        rightMargin,
        save,
      });
    },

    add_page_numbers: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const alignment = args.alignment || args.align || "center";
      const firstPage = args.firstPage !== undefined ? Boolean(args.firstPage) : false;
      const save = Boolean(args.save);

      return invokeOfficeCom("add_page_numbers", {
        path: targetPath,
        alignment,
        firstPage,
        save,
      });
    },

    insert_page_break: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const beforeText = args.beforeText || args.before || "";
      const afterText = args.afterText || args.after || "";
      const save = Boolean(args.save);

      return invokeOfficeCom("insert_page_break", {
        path: targetPath,
        beforeText,
        afterText,
        save,
      });
    },

    insert_table: async (args = {}) => {
      const app = args.app || "word";
      const targetPath = args.path || args.file || "";
      const rows = args.rows ? Number(args.rows) : 2;
      const cols = args.cols ? Number(args.cols) : 2;
      const data = Array.isArray(args.data) ? args.data : [];
      const headerBold = args.headerBold !== undefined ? Boolean(args.headerBold) : true;
      const headerColor = args.headerColor || null;
      const headerBgColor = args.headerBgColor || args.headerBg || null;
      const zebra = args.zebra !== undefined ? Boolean(args.zebra) : false;
      const zebraColor = args.zebraColor || null;
      const alignment = args.alignment || args.align || "center";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const left = args.left !== undefined ? Number(args.left) : null;
      const top = args.top !== undefined ? Number(args.top) : null;
      const width = args.width ? Number(args.width) : null;
      const height = args.height ? Number(args.height) : null;
      const save = Boolean(args.save);

      return invokeOfficeCom("insert_table", {
        app,
        path: targetPath,
        rows,
        cols,
        data,
        headerBold,
        headerColor,
        headerBgColor,
        zebra,
        zebraColor,
        alignment,
        slideIndex,
        left,
        top,
        width,
        height,
        save,
      });
    },

    word_table_format: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const tableIndex = args.tableIndex ? Number(args.tableIndex) : 1;
      const headerBg = args.headerBg || args.headerBackground || null;
      const zebra = args.zebra !== undefined ? Boolean(args.zebra) : true;
      const zebraBg = args.zebraBg || args.alternateBg || null;
      const borderColor = args.borderColor || null;
      const borderWidth = args.borderWidth ? Number(args.borderWidth) : null;
      const colWidths = args.colWidths || args.columnWidths || null;
      const save = Boolean(args.save);

      return invokeOfficeCom("word_table_format", {
        path: targetPath,
        tableIndex,
        headerBg,
        zebra,
        zebraBg,
        borderColor,
        borderWidth,
        colWidths,
        save,
      });
    },

    word_comment: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const searchText = args.searchText || args.search || args.target || "";
      const commentText = args.commentText || args.comment || args.text || "";
      const author = args.author || "Antigravity AI";
      const save = Boolean(args.save);

      if (!commentText) return { ok: false, error: "Parameter 'commentText' is required" };

      return invokeOfficeCom("word_comment", {
        path: targetPath,
        searchText,
        commentText,
        author,
        save,
      });
    },

    word_hyperlink: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const searchText = args.searchText || args.search || args.text || "";
      const url = args.url || args.address || args.href || "";
      const tooltip = args.tooltip || "";
      const save = Boolean(args.save);

      if (!searchText || !url) return { ok: false, error: "Parameters 'searchText' and 'url' are required" };

      return invokeOfficeCom("word_hyperlink", {
        path: targetPath,
        searchText,
        url,
        tooltip,
        save,
      });
    },

    word_template: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const presetName = args.preset || args.template || "academic_report";
      const themeName = args.theme || "politecnico_maroon";
      const title = args.title || "Document Title";
      const subtitle = args.subtitle || "";
      const author = args.author || "Author";
      const sections = Array.isArray(args.sections) ? args.sections : [];
      const mode = args.mode || "live";

      const theme = THEMES[themeName] || THEMES.politecnico_maroon;

      if (mode === "headless" && targetPath) {
        return headlessEngine.docxCreate(targetPath, {
          title,
          subtitle,
          author,
          sections,
          theme,
        });
      }

      if (targetPath) {
        const createRes = await invokeOfficeCom("create", { app: "word", path: targetPath, visible: true });
        if (!createRes.ok) return createRes;
      }

      await invokeOfficeCom("insert_cover_page", {
        path: targetPath,
        title,
        subtitle,
        author,
        primaryColor: theme.primary,
        secondaryColor: theme.secondary,
        save: false,
      });

      for (const sec of sections) {
        await invokeOfficeCom("append_section", {
          path: targetPath,
          heading: sec.heading || sec.title || "",
          body: sec.body || sec.content || "",
          headingLevel: sec.level || 1,
          headingColor: theme.primary,
          save: false,
        });
        if (sec.table && Array.isArray(sec.table)) {
          await invokeOfficeCom("insert_table", {
            path: targetPath,
            rows: sec.table.length,
            cols: sec.table[0]?.length || 2,
            data: sec.table,
            headerBgColor: theme.primary,
            save: false,
          });
        }
      }

      if (args.save) {
        await invokeOfficeCom("save", { path: targetPath, app: "word" });
      }

      return {
        ok: true,
        mode: "live_com",
        template: presetName,
        theme: themeName,
        path: targetPath,
        sectionsApplied: sections.length,
      };
    },

    insert_shape: async (args = {}) => {
      const app = args.app || "word";
      const targetPath = args.path || args.file || "";
      const shapeType = args.shapeType || args.type || "rectangle";
      const text = args.text !== undefined ? String(args.text) : (args.content !== undefined ? String(args.content) : "");
      const fontSize = args.fontSize || args.size ? Number(args.fontSize || args.size) : null;
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const textColor = args.textColor || args.color || null;
      const fillColor = args.fillColor || args.fill || null;
      const lineColor = args.lineColor || args.line || null;
      const lineWidth = args.lineWidth !== undefined ? Number(args.lineWidth) : null;
      const alignment = args.alignment || args.align || null;
      const textAlignment = args.textAlignment || null;
      const position = args.position || null;
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const left = args.left !== undefined ? Number(args.left) : null;
      const top = args.top !== undefined ? Number(args.top) : null;
      const width = args.width ? Number(args.width) : null;
      const height = args.height ? Number(args.height) : null;
      const save = Boolean(args.save);

      return invokeOfficeCom("insert_shape", {
        app,
        path: targetPath,
        shapeType,
        text,
        fontSize,
        bold,
        textColor,
        fillColor,
        lineColor,
        lineWidth,
        alignment,
        textAlignment,
        position,
        slideIndex,
        left,
        top,
        width,
        height,
        save,
      });
    },

    insert_image: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const imagePath = args.imagePath || args.image || "";
      const width = args.width ? Number(args.width) : null;
      const height = args.height ? Number(args.height) : null;
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const left = args.left ? Number(args.left) : null;
      const top = args.top ? Number(args.top) : null;
      const save = Boolean(args.save);

      if (!imagePath) return { ok: false, error: "Parameter 'imagePath' is required" };

      return invokeOfficeCom("insert_image", { path: targetPath, imagePath, width, height, slideIndex, left, top, save });
    },

    // ── 4. PowerPoint Actions ──────────────────────────────────────────────────
    read_slides: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      return invokeOfficeCom("read_slides", { path: targetPath });
    },

    add_slide: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const layout = args.layout || "title_content";
      const title = args.title || args.heading || "";
      const body = args.body || args.content || args.text || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : null;
      const titleColor = args.titleColor || null;
      const bodyColor = args.bodyColor || null;
      const save = Boolean(args.save);

      return invokeOfficeCom("add_slide", { path: targetPath, layout, title, body, slideIndex, titleColor, bodyColor, save });
    },

    set_slide_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const shapeIndex = args.shapeIndex ? Number(args.shapeIndex) : 1;
      const text = args.text !== undefined ? String(args.text) : (args.content !== undefined ? String(args.content) : "");
      const fontSize = args.fontSize || args.size ? Number(args.fontSize || args.size) : null;
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const color = args.color || null;
      const save = Boolean(args.save);

      return invokeOfficeCom("set_slide_text", { path: targetPath, slideIndex, shapeIndex, text, fontSize, bold, color, save });
    },

    delete_slide: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const save = Boolean(args.save);

      return invokeOfficeCom("delete_slide", { path: targetPath, slideIndex, save });
    },

    ppt_slide_ops: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const action = args.action || args.operation || "duplicate";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const targetIndex = args.targetIndex ? Number(args.targetIndex) : null;
      const width = args.width ? Number(args.width) : null;
      const height = args.height ? Number(args.height) : null;
      const orientation = args.orientation || null;
      const save = Boolean(args.save);

      return invokeOfficeCom("ppt_slide_ops", {
        path: targetPath,
        action,
        slideIndex,
        targetIndex,
        width,
        height,
        orientation,
        save,
      });
    },

    ppt_notes: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const notesText = args.notesText !== undefined ? String(args.notesText) : (args.notes !== undefined ? String(args.notes) : null);
      const action = notesText !== null ? "set" : "get";
      const save = Boolean(args.save);

      return invokeOfficeCom("ppt_notes", {
        path: targetPath,
        slideIndex,
        notesText: notesText || "",
        action,
        save,
      });
    },

    ppt_export_slide: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const outputPath = args.outputPath || args.output || "";
      const width = args.width ? Number(args.width) : 1920;
      const height = args.height ? Number(args.height) : 1080;

      if (!outputPath) return { ok: false, error: "Parameter 'outputPath' is required for ppt_export_slide" };

      return invokeOfficeCom("ppt_export_slide", {
        path: targetPath,
        slideIndex,
        outputPath,
        width,
        height,
      });
    },

    ppt_background: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const allSlides = Boolean(args.allSlides);
      const color = args.color || "#FFFFFF";
      const save = Boolean(args.save);

      return invokeOfficeCom("ppt_background", {
        path: targetPath,
        slideIndex,
        allSlides,
        color,
        save,
      });
    },

    ppt_chart: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const chartType = args.chartType || args.type || "column";
      const title = args.title || "";
      const left = args.left ? Number(args.left) : 100;
      const top = args.top ? Number(args.top) : 100;
      const width = args.width ? Number(args.width) : 500;
      const height = args.height ? Number(args.height) : 350;
      const save = Boolean(args.save);

      return invokeOfficeCom("ppt_chart", {
        path: targetPath,
        slideIndex,
        chartType,
        title,
        left,
        top,
        width,
        height,
        save,
      });
    },

    // ── 5. Excel Actions ───────────────────────────────────────────────────────
    read_range: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";

      if (args.mode === "headless" && targetPath && targetPath.toLowerCase().endsWith(".xlsx")) {
        return headlessEngine.excelRead(targetPath, { sheetName: sheet, maxRows: args.maxRows || 100 });
      }

      return invokeOfficeCom("read_range", { path: targetPath, sheet, range });
    },

    write_range: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const startCell = args.startCell || "A1";
      const data = Array.isArray(args.data) ? args.data : [];
      const headerBold = args.headerBold !== undefined ? Boolean(args.headerBold) : true;
      const headerColor = args.headerColor || null;
      const headerBgColor = args.headerBgColor || args.headerBg || null;
      const borders = Boolean(args.borders);
      const autoFit = Boolean(args.autoFit);
      const save = Boolean(args.save);

      if (!data.length) return { ok: false, error: "Parameter 'data' 2D array is required" };

      if (args.mode === "headless" && targetPath && targetPath.toLowerCase().endsWith(".xlsx")) {
        let headers = [];
        let rows = data;
        if (headerBold && data.length > 1) {
          headers = data[0];
          rows = data.slice(1);
        }
        return headlessEngine.excelWrite(targetPath, {
          sheet,
          startCell,
          headers,
          data: rows,
          headerBgColor,
          zebra: Boolean(args.zebra),
          borders,
          autoFit,
        });
      }

      return invokeOfficeCom("write_range", {
        path: targetPath,
        sheet,
        startCell,
        data,
        headerBold,
        headerColor,
        headerBgColor,
        borders,
        autoFit,
        save,
      });
    },

    write_cell: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const cell = args.cell || "";
      const row = args.row ? Number(args.row) : null;
      const col = args.col ? Number(args.col) : null;
      const value = args.value !== undefined ? args.value : (args.formula !== undefined ? args.formula : (args.text !== undefined ? args.text : ""));
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const fontSize = args.fontSize || args.size ? Number(args.fontSize || args.size) : null;
      const color = args.color || null;
      const bgColor = args.bgColor || args.fillColor || null;
      const numberFormat = args.numberFormat || null;
      const save = Boolean(args.save);

      return invokeOfficeCom("write_cell", {
        path: targetPath,
        sheet,
        cell,
        row,
        col,
        value,
        bold,
        fontSize,
        color,
        bgColor,
        numberFormat,
        save,
      });
    },

    add_sheet: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const name = args.name || args.sheetName || "";
      const save = Boolean(args.save);

      return invokeOfficeCom("add_sheet", { path: targetPath, name, save });
    },

    excel_table: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";
      const name = args.name || args.tableName || "";
      const style = args.style || "TableStyleMedium9";
      const showHeaders = args.showHeaders !== undefined ? Boolean(args.showHeaders) : true;
      const showTotals = Boolean(args.showTotals);
      const save = Boolean(args.save);

      if (!range) return { ok: false, error: "Parameter 'range' is required for excel_table" };

      return invokeOfficeCom("excel_table", {
        path: targetPath,
        sheet,
        range,
        name,
        style,
        showHeaders,
        showTotals,
        save,
      });
    },

    excel_chart: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";
      const chartType = args.chartType || args.type || "column";
      const title = args.title || "";
      const left = args.left ? Number(args.left) : 300;
      const top = args.top ? Number(args.top) : 20;
      const width = args.width ? Number(args.width) : 400;
      const height = args.height ? Number(args.height) : 250;
      const save = Boolean(args.save);

      if (!range) return { ok: false, error: "Parameter 'range' is required for excel_chart" };

      return invokeOfficeCom("excel_chart", {
        path: targetPath,
        sheet,
        range,
        chartType,
        title,
        left,
        top,
        width,
        height,
        save,
      });
    },

    excel_format_range: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";
      const numberFormat = args.numberFormat || args.format || null;
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const italic = args.italic !== undefined ? Boolean(args.italic) : null;
      const fontSize = args.fontSize ? Number(args.fontSize) : null;
      const fontColor = args.fontColor || args.color || null;
      const bgColor = args.bgColor || args.fillColor || null;
      const align = args.align || args.horizontalAlignment || null;
      const vAlign = args.vAlign || args.verticalAlignment || null;
      const wrapText = args.wrapText !== undefined ? Boolean(args.wrapText) : null;
      const borderStyle = args.borderStyle || null;
      const borderColor = args.borderColor || null;
      const rowHeight = args.rowHeight ? Number(args.rowHeight) : null;
      const colWidth = args.colWidth ? Number(args.colWidth) : null;
      const merge = args.merge !== undefined ? Boolean(args.merge) : null;
      const save = Boolean(args.save);

      if (!range) return { ok: false, error: "Parameter 'range' is required for excel_format_range" };

      return invokeOfficeCom("excel_format_range", {
        path: targetPath,
        sheet,
        range,
        numberFormat,
        bold,
        italic,
        fontSize,
        fontColor,
        bgColor,
        align,
        vAlign,
        wrapText,
        borderStyle,
        borderColor,
        rowHeight,
        colWidth,
        merge,
        save,
      });
    },

    excel_sheet_ops: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const action = args.action || args.operation || "rename";
      const newName = args.newName || "";
      const freezeCell = args.freezeCell || "A2";
      const save = Boolean(args.save);

      return invokeOfficeCom("excel_sheet_ops", {
        path: targetPath,
        sheet,
        action,
        newName,
        freezeCell,
        save,
      });
    },

    excel_calculate: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const scanErrors = args.scanErrors !== undefined ? Boolean(args.scanErrors) : true;
      const save = Boolean(args.save);
      return invokeOfficeCom("excel_calculate", { path: targetPath, scanErrors, save });
    },

    // ── 6. Batch Actions ───────────────────────────────────────────────────────
    batch_create: async (args = {}) => {
      const docs = Array.isArray(args.documents) ? args.documents : [];
      if (!docs.length) return { ok: false, error: "Array of 'documents' is required" };

      const results = [];
      for (const docSpec of docs) {
        const res = await invokeOfficeCom("create", {
          app: docSpec.app || "word",
          path: docSpec.path,
          template: docSpec.template || "",
          visible: docSpec.visible !== undefined ? Boolean(docSpec.visible) : false,
          overwrite: docSpec.overwrite !== undefined ? Boolean(docSpec.overwrite) : false,
        });
        results.push({ path: docSpec.path, result: res });
        if (!res.ok) {
          return { ok: false, error: `Batch creation stopped at ${docSpec.path}: ${res.error}`, results };
        }
      }

      return { ok: true, count: results.length, results };
    },

    batch_edit: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const operations = Array.isArray(args.operations) ? args.operations : [];

      if (!operations.length) return { ok: false, error: "Array of 'operations' is required" };

      const results = [];
      for (const op of operations) {
        const actionName = op.action || op.operation;
        if (!actionName || !actions[actionName]) {
          return { ok: false, error: `Unknown batch operation: ${actionName}`, results };
        }
        const opArgs = { ...(op.args || {}) };
        if (targetPath && !opArgs.path) opArgs.path = targetPath;
        const res = await actions[actionName](opArgs);
        results.push({ action: actionName, result: res });
        if (!res.ok) {
          return { ok: false, error: `Batch stopped on '${actionName}': ${res.error}`, results };
        }
      }

      return { ok: true, count: results.length, results };
    },

    capture_preview: async (args = {}) => {
      const app = args.app || "word";
      const outputPath = args.outputPath || args.output || "";
      const scrollToEnd = Boolean(args.scrollToEnd || args.scroll);
      const page = args.page || args.targetPage ? Number(args.page || args.targetPage) : null;
      const slide = args.slide || args.slideIndex ? Number(args.slide || args.slideIndex) : null;
      const sheet = args.sheet || null;
      return invokeOfficeCom("capture_preview", { app, outputPath, scrollToEnd, page, slide, sheet });
    },

    screen_capture: async (args = {}) => {
      return actions.capture_preview(args);
    },

    // ── Word Ribbon Suite Actions ─────────────────────────────────────────────
    apply_style: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const search = args.search || args.text || "";
      const style = args.style || "Heading 1";
      const save = Boolean(args.save);
      return invokeOfficeCom("apply_style", { path: targetPath, search, style, save });
    },

    insert_list: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const items = Array.isArray(args.items) ? args.items : (args.text ? [args.text] : []);
      const type = args.type || "bullet";
      const lineSpacing = args.lineSpacing ? Number(args.lineSpacing) : null;
      const spaceAfter = args.spaceAfter !== undefined ? Number(args.spaceAfter) : null;
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_list", { path: targetPath, items, type, lineSpacing, spaceAfter, save });
    },

    highlight_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const search = args.search || args.text || "";
      const color = args.color || "yellow";
      const save = Boolean(args.save);
      return invokeOfficeCom("highlight_text", { path: targetPath, search, color, save });
    },

    insert_callout: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const text = args.text || args.content || "";
      const title = args.title || "";
      const borderColor = args.borderColor || args.color || "#6C1D45";
      const bgColor = args.bgColor || args.fillColor || "#F7F9FA";
      const fontSize = args.fontSize ? Number(args.fontSize) : null;
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_callout", { path: targetPath, text, title, borderColor, bgColor, fontSize, save });
    },

    insert_divider: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const color = args.color || "#6C1D45";
      const weight = args.weight ? Number(args.weight) : 12;
      const spaceBefore = args.spaceBefore !== undefined ? Number(args.spaceBefore) : null;
      const spaceAfter = args.spaceAfter !== undefined ? Number(args.spaceAfter) : null;
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_divider", { path: targetPath, color, weight, spaceBefore, spaceAfter, save });
    },

    insert_header_footer: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const headerText = args.headerText || args.header || "";
      const footerText = args.footerText || args.footer || "";
      const headerAlign = args.headerAlign || "left";
      const footerAlign = args.footerAlign || "right";
      const differentFirstPage = args.differentFirstPage !== undefined ? Boolean(args.differentFirstPage) : true;
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_header_footer", { path: targetPath, headerText, footerText, headerAlign, footerAlign, differentFirstPage, save });
    },

    insert_cover_page: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const institution = args.institution || "INSTITUTO POLITÉCNICO NACIONAL";
      const school = args.school || "CENTRO DE ESTUDIOS CIENTÍFICOS Y TECNOLÓGICOS N° 8\n\"NARCISO BASSOLS\"";
      const subject = args.subject || "COMPUTACIÓN BÁSICA 1";
      const title = args.title || "INVESTIGACIÓN: CIBERESPACIO";
      const subtitle = args.subtitle || "Servicios de Internet, Identidad Digital, Ciberseguridad e Inteligencia Artificial";
      const group = args.group || "1IM12";
      const team = Array.isArray(args.team) ? args.team : (args.student ? [args.student] : ["Lara Lozano Mauricio Leonardo"]);
      const date = args.date || "16 de Octubre de 2024";
      const primaryColor = args.primaryColor || "#6C1D45";
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_cover_page", { path: targetPath, institution, school, subject, title, subtitle, group, team, date, primaryColor, save });
    },

    set_page_background: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const color = args.color || "#FFFFFF";
      const save = Boolean(args.save);
      return invokeOfficeCom("set_page_background", { path: targetPath, color, save });
    },

    insert_section_break: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const type = args.type || "next_page";
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_section_break", { path: targetPath, type, save });
    },

    insert_table_of_contents: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const pageBreak = args.pageBreak !== undefined ? Boolean(args.pageBreak) : true;
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_table_of_contents", { path: targetPath, pageBreak, save });
    },

    update_toc: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const save = Boolean(args.save);
      return invokeOfficeCom("update_toc", { path: targetPath, save });
    },

    insert_footnote: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const searchText = args.searchText || args.search || "";
      const footnoteText = args.footnoteText || args.text || "";
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_footnote", { path: targetPath, searchText, footnoteText, save });
    },

    replace_placeholders: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const replacements = args.replacements || {};
      const save = Boolean(args.save);
      return invokeOfficeCom("replace_placeholders", { path: targetPath, replacements, save });
    },

    get_document_stats: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      return invokeOfficeCom("get_document_stats", { path: targetPath });
    },

    set_language: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const languageId = args.languageId || 2058;
      const save = Boolean(args.save);
      return invokeOfficeCom("set_language", { path: targetPath, languageId, save });
    },

    set_view: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const viewType = args.viewType || "print";
      const zoomPercent = args.zoomPercent ? Number(args.zoomPercent) : 100;
      const showRuler = args.showRuler !== undefined ? Boolean(args.showRuler) : null;
      return invokeOfficeCom("set_view", { path: targetPath, viewType, zoomPercent, showRuler });
    },

    // ── 7. PDF & Headless Library Actions ─────────────────────────────────────
    pdf_inspect: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      if (!targetPath) return { ok: false, error: "Parameter 'path' is required for pdf_inspect" };
      return headlessEngine.pdfInspect(targetPath);
    },

    pdf_extract_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      if (!targetPath) return { ok: false, error: "Parameter 'path' is required for pdf_extract_text" };
      return headlessEngine.pdfExtractText(targetPath, {
        startPage: args.startPage,
        endPage: args.endPage,
        maxPages: args.maxPages,
      });
    },

    pdf_merge: async (args = {}) => {
      const files = args.files || args.paths || [];
      const outputPath = args.outputPath || args.output || "";
      if (!Array.isArray(files) || files.length < 2) {
        return { ok: false, error: "Parameter 'files' must be an array of at least 2 PDF paths" };
      }
      if (!outputPath) return { ok: false, error: "Parameter 'outputPath' is required for pdf_merge" };
      return headlessEngine.pdfMerge(files, outputPath);
    },

    get_templates: async (args = {}) => {
      return {
        ok: true,
        themes: THEMES,
        presets: TEMPLATE_PRESETS,
      };
    },

  };

  return {
    name: "documents",
    description: "Automatización integral y en tiempo real de Microsoft Office (Word, PowerPoint, Excel) y PDF en Windows con soporte LIVE y Headless.",
    permissions: {
      status: "low",
      list_open: "low",
      get_active: "low",
      inspect_structure: "low",
      create: "low",
      open: "low",
      save: "low",
      save_as: "low",
      close: "low",
      export_pdf: "low",
      read: "low",
      insert_text: "low",
      replace_text: "low",
      clear_document: "low",
      append_section: "low",
      format_text: "low",
      format_paragraphs: "low",
      page_setup: "low",
      add_page_numbers: "low",
      insert_page_break: "low",
      insert_table: "low",
      word_table_format: "low",
      word_comment: "low",
      word_hyperlink: "low",
      word_template: "low",
      insert_shape: "low",
      insert_image: "low",
      capture_preview: "low",
      screen_capture: "low",
      read_slides: "low",
      add_slide: "low",
      set_slide_text: "low",
      delete_slide: "low",
      ppt_slide_ops: "low",
      ppt_notes: "low",
      ppt_export_slide: "low",
      ppt_background: "low",
      ppt_chart: "low",
      read_range: "low",
      write_range: "low",
      write_cell: "low",
      add_sheet: "low",
      excel_table: "low",
      excel_chart: "low",
      excel_format_range: "low",
      excel_sheet_ops: "low",
      excel_calculate: "low",
      batch_create: "low",
      batch_edit: "low",
      apply_style: "low",
      insert_list: "low",
      highlight_text: "low",
      insert_callout: "low",
      insert_divider: "low",
      insert_header_footer: "low",
      insert_cover_page: "low",
      set_page_background: "low",
      insert_section_break: "low",
      insert_table_of_contents: "low",
      update_toc: "low",
      insert_footnote: "low",
      replace_placeholders: "low",
      get_document_stats: "low",
      set_language: "low",
      set_view: "low",
      pdf_inspect: "low",
      pdf_extract_text: "low",
      pdf_merge: "low",
      get_templates: "low",
    },
    actions
  };
}
