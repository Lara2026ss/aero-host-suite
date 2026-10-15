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

export const OFFICE_SHORTCUTS = {
  word: {
    general: [
      { key: "Ctrl + U / Ctrl + N", desc_es: "Crear nuevo documento", desc_en: "Create new document" },
      { key: "Ctrl + A / Ctrl + O", desc_es: "Abrir documento existente", desc_en: "Open document" },
      { key: "Ctrl + G / Ctrl + S", desc_es: "Guardar documento", desc_en: "Save document" },
      { key: "F12", desc_es: "Guardar como...", desc_en: "Save As" },
      { key: "Ctrl + P", desc_es: "Imprimir / Vista previa de impresión", desc_en: "Print / Print Preview" },
      { key: "Ctrl + R / Ctrl + W", desc_es: "Cerrar documento activo", desc_en: "Close active document" },
    ],
    history_and_clipboard: [
      { key: "Ctrl + Z", desc_es: "Deshacer la última acción", desc_en: "Undo last action" },
      { key: "Ctrl + Y", desc_es: "Rehacer la última acción deshecha", desc_en: "Redo last undone action" },
      { key: "Ctrl + C", desc_es: "Copiar texto o elemento seleccionado", desc_en: "Copy selection" },
      { key: "Ctrl + X", desc_es: "Cortar texto o elemento seleccionado", desc_en: "Cut selection" },
      { key: "Ctrl + V", desc_es: "Pegar contenido del portapapeles", desc_en: "Paste clipboard content" },
      { key: "Ctrl + Shift + V", desc_es: "Pegar solo texto / Pegar formato", desc_en: "Paste format / plain text" },
      { key: "Ctrl + E / Ctrl + A", desc_es: "Seleccionar todo el documento", desc_en: "Select all" },
    ],
    formatting: [
      { key: "Ctrl + N / Ctrl + B", desc_es: "Aplicar o quitar negrita", desc_en: "Bold toggle" },
      { key: "Ctrl + K / Ctrl + I", desc_es: "Aplicar o quitar cursiva", desc_en: "Italic toggle" },
      { key: "Ctrl + S / Ctrl + U", desc_es: "Aplicar o quitar subrayado", desc_en: "Underline toggle" },
      { key: "Shift + F3", desc_es: "Cambiar mayúsculas, minúsculas y tipo título", desc_en: "Change case (upper, lower, title)" },
      { key: "Ctrl + Barra Espaciadora", desc_es: "Borrar todo el formato del texto seleccionado", desc_en: "Clear formatting" },
      { key: "Ctrl + Shift + >", desc_es: "Aumentar tamaño de fuente", desc_en: "Increase font size" },
      { key: "Ctrl + Shift + <", desc_es: "Disminuir tamaño de fuente", desc_en: "Decrease font size" },
    ],
    alignment_and_spacing: [
      { key: "Ctrl + T / Ctrl + E", desc_es: "Centrar párrafo", desc_en: "Center alignment" },
      { key: "Ctrl + Q / Ctrl + L", desc_es: "Alinear párrafo a la izquierda", desc_en: "Align left" },
      { key: "Ctrl + D / Ctrl + R", desc_es: "Alinear párrafo a la derecha", desc_en: "Align right" },
      { key: "Ctrl + J", desc_es: "Justificar párrafo", desc_en: "Justify paragraph" },
      { key: "Ctrl + 1", desc_es: "Interlineado simple (1.0)", desc_en: "Single line spacing" },
      { key: "Ctrl + 2", desc_es: "Interlineado doble (2.0)", desc_en: "Double line spacing" },
      { key: "Ctrl + 5", desc_es: "Interlineado 1.5", desc_en: "1.5 line spacing" },
    ],
    navigation_and_insert: [
      { key: "Ctrl + Enter", desc_es: "Insertar salto de página", desc_en: "Insert page break" },
      { key: "Ctrl + Alt + K", desc_es: "Insertar hipervínculo", desc_en: "Insert hyperlink" },
      { key: "Ctrl + Alt + L", desc_es: "Insertar viñetas automáticas", desc_en: "Insert bullet list" },
      { key: "Alt + Ctrl + F", desc_es: "Insertar nota al pie", desc_en: "Insert footnote" },
      { key: "Ctrl + B / Ctrl + F", desc_es: "Buscar en el documento", desc_en: "Find in document" },
      { key: "Ctrl + L / Ctrl + H", desc_es: "Reemplazar texto en el documento", desc_en: "Replace text" },
      { key: "Ctrl + Shift + E", desc_es: "Activar o desactivar Control de Cambios", desc_en: "Track changes toggle" },
    ]
  },
  excel: {
    general_and_navigation: [
      { key: "Ctrl + U / Ctrl + N", desc_es: "Nuevo libro de trabajo", desc_en: "New workbook" },
      { key: "Ctrl + G / Ctrl + S", desc_es: "Guardar libro", desc_en: "Save workbook" },
      { key: "F12", desc_es: "Guardar como...", desc_en: "Save As" },
      { key: "Ctrl + Z", desc_es: "Deshacer última acción", desc_en: "Undo last action" },
      { key: "Ctrl + Y", desc_es: "Rehacer última acción", desc_en: "Redo last action" },
      { key: "F4", desc_es: "Repetir última acción / Alternar referencias relativas/absolutas ($A$1)", desc_en: "Repeat action / Toggle absolute reference" },
      { key: "Ctrl + RePág / AvPág", desc_es: "Moverse entre hojas de cálculo", desc_en: "Switch between worksheets" },
    ],
    formatting_and_selection: [
      { key: "Ctrl + 1", desc_es: "Abrir cuadro de diálogo Formato de Celdas", desc_en: "Format Cells dialog" },
      { key: "Ctrl + Shift + ~", desc_es: "Aplicar formato General de número", desc_en: "General number format" },
      { key: "Ctrl + Shift + 4 ($)", desc_es: "Aplicar formato Moneda ($)", desc_en: "Currency format" },
      { key: "Ctrl + Shift + 5 (%)", desc_es: "Aplicar formato Porcentaje (%)", desc_en: "Percentage format" },
      { key: "Ctrl + Espacio", desc_es: "Seleccionar columna completa", desc_en: "Select entire column" },
      { key: "Shift + Espacio", desc_es: "Seleccionar fila completa", desc_en: "Select entire row" },
      { key: "Ctrl + +", desc_es: "Insertar nuevas celdas, filas o columnas", desc_en: "Insert cells, rows, or columns" },
      { key: "Ctrl + -", desc_es: "Eliminar celdas, filas o columnas seleccionadas", desc_en: "Delete cells, rows, or columns" },
    ],
    data_and_formulas: [
      { key: "Ctrl + Shift + L", desc_es: "Activar o desactivar AutoFiltro en la tabla/rango", desc_en: "Toggle AutoFilter" },
      { key: "Alt + Flecha Abajo", desc_es: "Desplegar menú de filtro o lista de validación de datos", desc_en: "Open filter/validation dropdown" },
      { key: "Ctrl + T", desc_es: "Crear tabla oficial estructurada (ListObject)", desc_en: "Create Table" },
      { key: "Alt + =", desc_es: "Insertar fórmula Autosuma (=SUMA())", desc_en: "AutoSum formula" },
      { key: "F9", desc_es: "Recalcular todas las fórmulas de todos los libros abiertos", desc_en: "Calculate all formulas" },
      { key: "Shift + F9", desc_es: "Recalcular la hoja de cálculo activa", desc_en: "Calculate active sheet" },
      { key: "Shift + F3", desc_es: "Insertar función con el asistente", desc_en: "Insert function wizard" },
    ]
  },
  powerpoint: {
    presentation: [
      { key: "F5", desc_es: "Iniciar presentación con diapositivas desde el principio", desc_en: "Start slide show from beginning" },
      { key: "Shift + F5", desc_es: "Iniciar presentación desde la diapositiva actual", desc_en: "Start slide show from current slide" },
      { key: "Esc", desc_es: "Finalizar presentación de diapositivas", desc_en: "End slide show" },
      { key: "B", desc_es: "Pantalla en negro durante la presentación", desc_en: "Black screen during show" },
      { key: "W", desc_es: "Pantalla en blanco durante la presentación", desc_en: "White screen during show" },
    ],
    slide_and_object_design: [
      { key: "Ctrl + M", desc_es: "Insertar nueva diapositiva", desc_en: "Insert new slide" },
      { key: "Ctrl + D", desc_es: "Duplicar diapositiva u objeto seleccionado", desc_en: "Duplicate slide or selected object" },
      { key: "Ctrl + G", desc_es: "Agrupar formas u objetos seleccionados", desc_en: "Group selected objects" },
      { key: "Ctrl + Shift + G", desc_es: "Desagrupar formas u objetos", desc_en: "Ungroup objects" },
      { key: "Ctrl + Z", desc_es: "Deshacer última acción", desc_en: "Undo" },
      { key: "Ctrl + Y", desc_es: "Rehacer última acción", desc_en: "Redo" },
    ]
  }
};

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

    undo: async (args = {}) => {
      const app = args.app || "word";
      const times = args.times ? Number(args.times) : 1;
      const targetPath = args.path || args.file || "";
      return invokeOfficeCom("undo", { app, times, path: targetPath });
    },

    redo: async (args = {}) => {
      const app = args.app || "word";
      const times = args.times ? Number(args.times) : 1;
      const targetPath = args.path || args.file || "";
      return invokeOfficeCom("redo", { app, times, path: targetPath });
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

    compose_document: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const title = args.title || "Documento";
      const subtitle = args.subtitle || "";
      const author = args.author || "Autor";
      const institution = args.institution || "";
      const school = args.school || "";
      const subject = args.subject || "";
      const group = args.group || "";
      const date = args.date || new Date().toLocaleDateString("es-ES", { year: "numeric", month: "long", day: "numeric" });
      const coverPage = args.coverPage !== undefined ? Boolean(args.coverPage) : false;
      const themeName = args.theme || "modern_blue";
      const sections = Array.isArray(args.sections) ? args.sections : [];
      const toc = Boolean(args.toc || args.tableOfContents);
      const pageNumbers = args.pageNumbers !== undefined ? Boolean(args.pageNumbers) : true;
      const watermark = args.watermark || null;
      const mode = args.mode || "live";
      const save = args.save !== undefined ? Boolean(args.save) : true;

      const theme = THEMES[themeName] || THEMES.modern_blue;

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

      if (coverPage) {
        await invokeOfficeCom("insert_cover_page", {
          path: targetPath,
          institution: institution || "INSTITUTO POLITÉCNICO NACIONAL",
          school: school || "CENTRO DE ESTUDIOS CIENTÍFICOS Y TECNOLÓGICOS",
          subject: subject || "Investigación",
          title,
          subtitle,
          group,
          team: Array.isArray(author) ? author : [author],
          date,
          primaryColor: theme.primary,
          save: false,
        });
      }

      if (watermark) {
        await invokeOfficeCom("insert_watermark", {
          path: targetPath,
          text: typeof watermark === "string" ? watermark : (watermark.text || "CONFIDENCIAL"),
          color: watermark.color || "#E8E8E8",
          save: false,
        });
      }

      if (toc) {
        await invokeOfficeCom("insert_table_of_contents", {
          path: targetPath,
          pageBreak: true,
          save: false,
        });
      }

      for (const sec of sections) {
        const heading = sec.heading || sec.title || "";
        const level = sec.level ? Number(sec.level) : 1;
        const body = sec.body || sec.content || sec.text || "";
        const paragraphs = Array.isArray(sec.paragraphs) ? sec.paragraphs : (body ? [body] : []);

        if (heading) {
          await invokeOfficeCom("append_section", {
            path: targetPath,
            heading,
            headingLevel: level,
            headingColor: level === 1 ? theme.primary : theme.secondary,
            pageBreak: Boolean(sec.pageBreak),
            save: false,
          });
        }

        for (const p of paragraphs) {
          if (p) {
            await invokeOfficeCom("insert_text", {
              path: targetPath,
              text: p,
              position: "end",
              fontSize: 11,
              fontName: theme.fontFamily || "Calibri",
              lineSpacing: 1.15,
              spaceAfter: 6,
              save: false,
            });
          }
        }

        if (sec.callout) {
          const cText = typeof sec.callout === "string" ? sec.callout : (sec.callout.text || "");
          const cTitle = typeof sec.callout === "object" ? sec.callout.title : "";
          if (cText) {
            await invokeOfficeCom("insert_callout", {
              path: targetPath,
              text: cText,
              title: cTitle,
              borderColor: theme.primary,
              bgColor: theme.background || "#F7F9FA",
              save: false,
            });
          }
        }

        if (Array.isArray(sec.list) && sec.list.length > 0) {
          await invokeOfficeCom("insert_list", {
            path: targetPath,
            items: sec.list,
            type: sec.listType || "bullet",
            save: false,
          });
        }

        if (Array.isArray(sec.table) && sec.table.length > 0) {
          await invokeOfficeCom("insert_table", {
            path: targetPath,
            rows: sec.table.length,
            cols: sec.table[0]?.length || 2,
            data: sec.table,
            headerBgColor: theme.primary,
            zebra: true,
            zebraColor: "#F9FAFB",
            save: false,
          });
        }

        if (sec.divider) {
          await invokeOfficeCom("insert_divider", {
            path: targetPath,
            color: theme.accent || theme.primary,
            weight: 12,
            save: false,
          });
        }
      }

      if (pageNumbers) {
        await invokeOfficeCom("add_page_numbers", {
          path: targetPath,
          alignment: "right",
          firstPage: !coverPage,
          save: false,
        });
      }

      if (args.header || args.footer) {
        await invokeOfficeCom("insert_header_footer", {
          path: targetPath,
          headerText: args.header || "",
          footerText: args.footer || "",
          differentFirstPage: coverPage,
          save: false,
        });
      }

      if (save) {
        await invokeOfficeCom("save", { path: targetPath, app: "word" });
      }

      return {
        ok: true,
        mode: "live_com",
        composed: true,
        title,
        sectionsCount: sections.length,
        theme: themeName,
        path: targetPath,
      };
    },

    change_case: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const caseType = args.case || args.type || "upper";
      const search = args.search || "";
      const paragraph = args.paragraph ? Number(args.paragraph) : null;
      const save = Boolean(args.save);
      return invokeOfficeCom("change_case", { path: targetPath, case: caseType, search, paragraph, save });
    },

    clear_formatting: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const paragraph = args.paragraph ? Number(args.paragraph) : null;
      const all = Boolean(args.all);
      const save = Boolean(args.save);
      return invokeOfficeCom("clear_formatting", { path: targetPath, paragraph, all, save });
    },

    paragraph_indent: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const action = args.action || "increase";
      const paragraph = args.paragraph ? Number(args.paragraph) : null;
      const leftIndent = args.leftIndent ? Number(args.leftIndent) : null;
      const rightIndent = args.rightIndent ? Number(args.rightIndent) : null;
      const firstLineIndent = args.firstLineIndent ? Number(args.firstLineIndent) : null;
      const save = Boolean(args.save);
      return invokeOfficeCom("paragraph_indent", { path: targetPath, action, paragraph, leftIndent, rightIndent, firstLineIndent, save });
    },

    paragraph_shading: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const paragraph = args.paragraph ? Number(args.paragraph) : 1;
      const color = args.color || args.bgColor || "#F2F4F7";
      const save = Boolean(args.save);
      return invokeOfficeCom("paragraph_shading", { path: targetPath, paragraph, color, save });
    },

    insert_textbox: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const text = args.text || args.content || "Text Box";
      const left = args.left ? Number(args.left) : 100;
      const top = args.top ? Number(args.top) : 100;
      const width = args.width ? Number(args.width) : 250;
      const height = args.height ? Number(args.height) : 80;
      const fontSize = args.fontSize ? Number(args.fontSize) : null;
      const fontName = args.fontName || null;
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const color = args.color || null;
      const fillColor = args.fillColor || null;
      const borderColor = args.borderColor || null;
      const borderWidth = args.borderWidth ? Number(args.borderWidth) : null;
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_textbox", { path: targetPath, text, left, top, width, height, fontSize, fontName, bold, color, fillColor, borderColor, borderWidth, save });
    },

    insert_watermark: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const text = args.text || "CONFIDENCIAL";
      const fontName = args.fontName || "Calibri";
      const fontSize = args.fontSize ? Number(args.fontSize) : 72;
      const color = args.color || "#D9D9D9";
      const diagonal = args.diagonal !== undefined ? Boolean(args.diagonal) : true;
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_watermark", { path: targetPath, text, fontName, fontSize, color, diagonal, save });
    },

    set_columns: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const columns = args.columns ? Number(args.columns) : 2;
      const spacing = args.spacing ? Number(args.spacing) : 36;
      const lineBetween = Boolean(args.lineBetween);
      const section = args.section ? Number(args.section) : 1;
      const save = Boolean(args.save);
      return invokeOfficeCom("set_columns", { path: targetPath, columns, spacing, lineBetween, section, save });
    },

    track_changes: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const action = args.action || "toggle";
      const save = Boolean(args.save);
      return invokeOfficeCom("track_changes", { path: targetPath, action, save });
    },

    insert_symbol: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const symbol = args.symbol || args.char || "©";
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_symbol", { path: targetPath, symbol, save });
    },

    insert_caption: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const label = args.label || "Figura";
      const title = args.title || "Descripción";
      const position = args.position || "below";
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_caption", { path: targetPath, label, title, position, save });
    },

    insert_bookmark: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const name = args.name || "Bookmark1";
      const save = Boolean(args.save);
      return invokeOfficeCom("insert_bookmark", { path: targetPath, name, save });
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

    ppt_transition: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const effect = args.effect || "fade";
      const duration = args.duration ? Number(args.duration) : null;
      const all = Boolean(args.all);
      const save = Boolean(args.save);

      return invokeOfficeCom("ppt_transition", {
        path: targetPath,
        slideIndex,
        effect,
        duration,
        all,
        save,
      });
    },

    ppt_align_shapes: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const align = args.align || "center";
      const save = Boolean(args.save);

      return invokeOfficeCom("ppt_align_shapes", {
        path: targetPath,
        slideIndex,
        align,
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

    excel_auto_filter: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";
      const field = args.field ? Number(args.field) : null;
      const criteria = args.criteria || "";
      const save = Boolean(args.save);
      return invokeOfficeCom("excel_auto_filter", { path: targetPath, sheet, range, field, criteria, save });
    },

    excel_data_validation: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";
      const list = args.list || ["Sí", "No", "Pendiente"];
      const title = args.title || "";
      const message = args.message || "";
      const save = Boolean(args.save);
      return invokeOfficeCom("excel_data_validation", { path: targetPath, sheet, range, list, title, message, save });
    },

    excel_sort: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";
      const keyCell = args.keyCell || "";
      const descending = Boolean(args.descending);
      const hasHeader = args.hasHeader !== undefined ? Boolean(args.hasHeader) : true;
      const save = Boolean(args.save);
      return invokeOfficeCom("excel_sort", { path: targetPath, sheet, range, keyCell, descending, hasHeader, save });
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

    get_shortcuts: async (args = {}) => {
      const app = (args.app || "all").toLowerCase();
      const category = (args.category || "").toLowerCase();

      if (app !== "all" && OFFICE_SHORTCUTS[app]) {
        const shortcuts = OFFICE_SHORTCUTS[app];
        if (category && shortcuts[category]) {
          return {
            ok: true,
            app,
            category,
            shortcuts: shortcuts[category],
          };
        }
        return {
          ok: true,
          app,
          categories: Object.keys(shortcuts),
          shortcuts,
        };
      }

      return {
        ok: true,
        apps: Object.keys(OFFICE_SHORTCUTS),
        shortcuts: OFFICE_SHORTCUTS,
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
      undo: "low",
      redo: "low",
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
      compose_document: "low",
      change_case: "low",
      clear_formatting: "low",
      paragraph_indent: "low",
      paragraph_shading: "low",
      insert_textbox: "low",
      insert_watermark: "low",
      set_columns: "low",
      track_changes: "low",
      insert_symbol: "low",
      insert_caption: "low",
      insert_bookmark: "low",
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
      ppt_transition: "low",
      ppt_align_shapes: "low",
      read_range: "low",
      write_range: "low",
      write_cell: "low",
      add_sheet: "low",
      excel_table: "low",
      excel_chart: "low",
      excel_format_range: "low",
      excel_sheet_ops: "low",
      excel_calculate: "low",
      excel_auto_filter: "low",
      excel_data_validation: "low",
      excel_sort: "low",
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
      get_shortcuts: "low",
    },
    actions
  };
}
