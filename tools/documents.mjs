/**
 * FLUXER XZ — tools/documents.mjs
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

      if (!targetPath) {
        return { ok: false, error: "Parameter 'path' is required for creating a document" };
      }

      return invokeOfficeCom("create", { app, path: targetPath, template, visible, overwrite });
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

      if (!targetPath) {
        return { ok: false, error: "Parameter 'path' is required for read" };
      }

      return invokeOfficeCom("read", { path: targetPath, maxChars });
    },

    insert_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const text = args.text || args.content || "";
      const position = args.position || "end";
      const bold = args.bold !== undefined ? Boolean(args.bold) : null;
      const italic = args.italic !== undefined ? Boolean(args.italic) : null;
      const fontSize = args.fontSize ? Number(args.fontSize) : null;

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };
      if (!text) return { ok: false, error: "Parameter 'text' is required" };

      return invokeOfficeCom("insert_text", { path: targetPath, text, position, bold, italic, fontSize });
    },

    replace_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const search = args.search || args.find || "";
      const replace = args.replace !== undefined ? String(args.replace) : "";
      const matchCase = Boolean(args.matchCase);
      const matchWholeWord = Boolean(args.matchWholeWord);

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };
      if (!search) return { ok: false, error: "Parameter 'search' is required" };

      return invokeOfficeCom("replace_text", { path: targetPath, search, replace, matchCase, matchWholeWord });
    },

    append_section: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const heading = args.heading || args.title || "";
      const body = args.body || args.content || "";
      const headingLevel = args.headingLevel ? Number(args.headingLevel) : 1;

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };
      if (!heading) return { ok: false, error: "Parameter 'heading' is required" };

      return invokeOfficeCom("append_section", { path: targetPath, heading, body, headingLevel });
    },

    insert_table: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const rows = args.rows ? Number(args.rows) : 2;
      const cols = args.cols ? Number(args.cols) : 2;
      const data = Array.isArray(args.data) ? args.data : [];

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };

      return invokeOfficeCom("insert_table", { path: targetPath, rows, cols, data });
    },

    insert_image: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const imagePath = args.imagePath || args.image || "";
      const width = args.width ? Number(args.width) : null;
      const height = args.height ? Number(args.height) : null;
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const left = args.left ? Number(args.left) : null;
      const top = args.top ? Number(args.top) : null;

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };
      if (!imagePath) return { ok: false, error: "Parameter 'imagePath' is required" };

      return invokeOfficeCom("insert_image", { path: targetPath, imagePath, width, height, slideIndex, left, top });
    },

    // ── 4. PowerPoint Actions ──────────────────────────────────────────────────
    read_slides: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };

      return invokeOfficeCom("read_slides", { path: targetPath });
    },

    add_slide: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const layout = args.layout || "title_content";
      const title = args.title || "";
      const body = args.body || "";

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };

      return invokeOfficeCom("add_slide", { path: targetPath, layout, title, body });
    },

    set_slide_text: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;
      const shapeIndex = args.shapeIndex ? Number(args.shapeIndex) : 1;
      const text = args.text !== undefined ? String(args.text) : "";

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };

      return invokeOfficeCom("set_slide_text", { path: targetPath, slideIndex, shapeIndex, text });
    },

    delete_slide: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const slideIndex = args.slideIndex ? Number(args.slideIndex) : 1;

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };

      return invokeOfficeCom("delete_slide", { path: targetPath, slideIndex });
    },

    // ── 5. Excel Actions ───────────────────────────────────────────────────────
    read_range: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const range = args.range || "";

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };

      return invokeOfficeCom("read_range", { path: targetPath, sheet, range });
    },

    write_range: async (args = {}) => {
      const targetPath = args.path || args.file || "";
      const sheet = args.sheet || "";
      const startCell = args.startCell || "A1";
      const data = Array.isArray(args.data) ? args.data : [];

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };
      if (!data.length) return { ok: false, error: "Parameter 'data' array is required" };

      return invokeOfficeCom("write_range", { path: targetPath, sheet, startCell, data });
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

      if (!targetPath) return { ok: false, error: "Parameter 'path' is required" };
      if (!operations.length) return { ok: false, error: "Array of 'operations' is required" };

      const results = [];
      for (const op of operations) {
        const actionName = op.action || op.operation;
        if (!actionName || !actions[actionName]) {
          return { ok: false, error: `Unknown batch operation: ${actionName}`, results };
        }
        const res = await actions[actionName]({ path: targetPath, ...(op.args || {}) });
        results.push({ action: actionName, result: res });
        if (!res.ok) {
          return { ok: false, error: `Batch stopped on '${actionName}': ${res.error}`, results };
        }
      }

      return { ok: true, count: results.length, results };
    },
  };

  return {
    name: "documents",
    description: "Automatización integral y en tiempo real de Microsoft Office (Word, PowerPoint, Excel) y PDF en Windows con soporte LIVE y Headless.",
    permissions: {
      status: "low",
      list_open: "low",
      get_active: "low",
      create: "medium",
      open: "low",
      save: "medium",
      save_as: "medium",
      close: "medium",
      export_pdf: "medium",
      read: "low",
      insert_text: "medium",
      replace_text: "medium",
      append_section: "medium",
      insert_table: "medium",
      insert_image: "medium",
      read_slides: "low",
      add_slide: "medium",
      set_slide_text: "medium",
      delete_slide: "destructive",
      read_range: "low",
      write_range: "medium",
      batch_create: "medium",
      batch_edit: "medium",
    },
    actions
  };
}
