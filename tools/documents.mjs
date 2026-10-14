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
        return JSON.parse(raw);
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

    // ── 3. Word Actions ────────────────────────────────────────────────────────
    read: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const maxChars = args.maxChars ? Number(args.maxChars) : 4000;
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

    // ── 5. Excel Actions ───────────────────────────────────────────────────────
    read_range: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";

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

  };

  return {
    name: "documents",
    description: "Automatización integral y en tiempo real de Microsoft Office (Word, PowerPoint, Excel) y PDF en Windows con soporte LIVE y Headless.",
    permissions: {
      status: "low",
      list_open: "low",
      get_active: "low",
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
      insert_shape: "low",
      insert_image: "low",
      capture_preview: "low",
      screen_capture: "low",
      read_slides: "low",
      add_slide: "low",
      set_slide_text: "low",
      delete_slide: "low",
      read_range: "low",
      write_range: "low",
      write_cell: "low",
      add_sheet: "low",
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
    },
    actions
  };
}
