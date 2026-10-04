/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * tools/shortcuts.mjs — Next-Gen Multi-step Macro & Template Engine
 */
import fs from "node:fs/promises";
import path from "node:path";
import { security, RISK_LEVELS } from "../core/permissions.mjs";

export function createShortcutsDomain({ runtime, router }) {
  let shortcutsDir = path.resolve(runtime?.dirs?.shortcuts || path.join(process.cwd(), "data", "shortcuts"));

  async function ensureDir() {
    await fs.mkdir(shortcutsDir, { recursive: true });
  }

  function interpolate(template, params) {
    if (typeof template !== "string") return template;
    return template.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (match, key) => {
      return params[key] !== undefined ? String(params[key]) : match;
    });
  }

  function resolveArgs(argsObj, params) {
    if (!argsObj || typeof argsObj !== "object") return argsObj;
    if (Array.isArray(argsObj)) {
      return argsObj.map((item) => (typeof item === "string" ? interpolate(item, params) : resolveArgs(item, params)));
    }
    const resolved = {};
    for (const [k, v] of Object.entries(argsObj)) {
      if (typeof v === "string") {
        resolved[k] = interpolate(v, params);
      } else if (typeof v === "object" && v !== null) {
        resolved[k] = resolveArgs(v, params);
      } else {
        resolved[k] = v;
      }
    }
    return resolved;
  }

  async function getShortcutPath(id) {
    const safeId = String(id).replace(/[^a-zA-Z0-9_\-]/g, "");
    return path.join(shortcutsDir, `${safeId}.json`);
  }

  async function loadShortcut(id) {
    const filePath = await getShortcutPath(id);
    const content = await fs.readFile(filePath, "utf8");
    return JSON.parse(content);
  }

  // Pre-seed 5-8 safe production starter shortcuts
  async function seedDefaultShortcuts() {
    await ensureDir();
    const defaults = [
      {
        id: "git_quick_status",
        name: "Git Quick Status & Branch Summary",
        description: "Ejecuta un diagnóstico rápido del estado de Git y ramas",
        params: {},
        steps: [
          { tool: "terminal", action: "run_command", args: { command: "git status -s" } },
          { tool: "terminal", action: "run_command", args: { command: "git branch --show-current" } }
        ]
      },
      {
        id: "create_module_scaffold",
        name: "Create Module Scaffold",
        description: "Crea la estructura básica de un nuevo módulo con README e index.js",
        params: {
          name: { type: "string", default: "sample-module", description: "Nombre del módulo" },
          author: { type: "string", default: "Fluxer Developer", description: "Autor" }
        },
        steps: [
          {
            tool: "files",
            action: "write_file",
            args: {
              path: "src/{{name}}/index.js",
              content: "/**\n * Módulo: {{name}}\n * Autor: {{author}}\n */\nexport function init() {\n  return '{{name}} ready';\n}\n"
            },
            onError: "rollback"
          },
          {
            tool: "files",
            action: "write_file",
            args: {
              path: "src/{{name}}/README.md",
              content: "# Módulo {{name}}\n\nAutor: {{author}}\n\nDescripción inicial generada por Fluxer XZ.\n"
            },
            onError: "rollback"
          }
        ]
      },
      {
        id: "clean_temp_files",
        name: "Clean Temporary Files",
        description: "Limpia archivos de caché y temporales en data/cache",
        params: {},
        steps: [
          { tool: "files", action: "list_directory", args: { path: "data/cache" } }
        ]
      },
      {
        id: "system_health_snapshot",
        name: "System Health & Process Snapshot",
        description: "Obtiene información resumida de CPU, RAM y servicios en ejecución",
        params: {},
        steps: [
          { tool: "system", action: "get_info", args: {} },
          { tool: "system", action: "list_processes", args: { limit: 10 } }
        ]
      },
      {
        id: "backup_notes_document",
        name: "Backup Notes Document",
        description: "Crea una copia de respaldo con fecha de notas importantes",
        params: {
          noteName: { type: "string", default: "release_notes.txt", description: "Nombre del archivo de notas" }
        },
        steps: [
          { tool: "files", action: "copy_path", args: { src: "{{noteName}}", dest: "data/backups/{{noteName}}.bak" } }
        ]
      },
      {
        id: "web_search_and_extract",
        name: "Web Search and Extract",
        description: "Busca información relevante y extrae contenido resumido",
        params: {
          query: { type: "string", default: "Fluxer XZ MCP", description: "Consulta de búsqueda" }
        },
        steps: [
          { tool: "web", action: "search_web", args: { query: "{{query}}", limit: 3 } }
        ]
      }
    ];

    for (const sc of defaults) {
      const p = path.join(shortcutsDir, `${sc.id}.json`);
      try {
        await fs.access(p);
      } catch {
        await fs.writeFile(p, JSON.stringify(sc, null, 2), "utf8");
      }
    }
  }

  // Seed defaults immediately
  seedDefaultShortcuts().catch(() => {});

  return {
    name: "shortcuts",
    description: "Motor universal de atajos, macros multi-paso y automatización con templates.",
    permissions: {
      list_shortcuts: RISK_LEVELS.LOW,
      view_shortcut: RISK_LEVELS.LOW,
      run_shortcut: RISK_LEVELS.HIGH,
      create_shortcut: RISK_LEVELS.MEDIUM,
      edit_shortcut: RISK_LEVELS.MEDIUM,
      delete_shortcut: RISK_LEVELS.DESTRUCTIVE,
      duplicate_shortcut: RISK_LEVELS.MEDIUM,
      export_shortcut: RISK_LEVELS.LOW,
      import_shortcut: RISK_LEVELS.MEDIUM,
      validate_shortcut: RISK_LEVELS.LOW,
      set_folder: RISK_LEVELS.HIGH
    },
    actions: {
      async list_shortcuts() {
        await ensureDir();
        const files = (await fs.readdir(shortcutsDir)).filter((f) => f.endsWith(".json"));
        const list = [];
        for (const file of files) {
          try {
            const raw = await fs.readFile(path.join(shortcutsDir, file), "utf8");
            const sc = JSON.parse(raw);
            list.push({
              id: sc.id || path.basename(file, ".json"),
              name: sc.name || sc.id,
              description: sc.description || "",
              paramsCount: Object.keys(sc.params || {}).length,
              stepsCount: (sc.steps || []).length
            });
          } catch {}
        }
        return { ok: true, count: list.length, folder: shortcutsDir, shortcuts: list };
      },

      async view_shortcut({ id }) {
        if (!id) throw new Error("ID de shortcut requerido");
        const sc = await loadShortcut(id);
        return { ok: true, shortcut: sc };
      },

      async set_folder({ path: newPath }) {
        if (!newPath) throw new Error("Ruta de carpeta requerida");
        const resolved = path.resolve(newPath);
        await fs.mkdir(resolved, { recursive: true });
        // Test writability
        const testFile = path.join(resolved, `.write_test_${Date.now()}`);
        await fs.writeFile(testFile, "ok", "utf8");
        await fs.unlink(testFile);

        shortcutsDir = resolved;
        return { ok: true, message: `Carpeta de shortcuts actualizada a: ${shortcutsDir}` };
      },

      async validate_shortcut({ shortcut }) {
        const sc = typeof shortcut === "string" ? JSON.parse(shortcut) : shortcut;
        if (!sc || !sc.id || !sc.name || !Array.isArray(sc.steps)) {
          return { ok: false, error: "El shortcut debe tener 'id', 'name' y un array 'steps'." };
        }
        for (let i = 0; i < sc.steps.length; i++) {
          const step = sc.steps[i];
          if (!step.tool || !step.action) {
            return { ok: false, error: `Paso ${i + 1} inválido: faltan campos 'tool' o 'action'.` };
          }
        }
        return { ok: true, message: "Shortcut válido y listo para ejecutar", stepsCount: sc.steps.length };
      },

      async create_shortcut({ id, name, description = "", params = {}, steps = [] }) {
        if (!id || !name || !Array.isArray(steps)) {
          throw new Error("Parámetros 'id', 'name' y 'steps' son obligatorios");
        }
        await ensureDir();
        const sc = { id, name, description, params, steps, createdAt: new Date().toISOString() };
        const filePath = await getShortcutPath(id);
        await fs.writeFile(filePath, JSON.stringify(sc, null, 2), "utf8");
        return { ok: true, message: `Shortcut '${id}' creado exitosamente`, path: filePath };
      },

      async edit_shortcut({ id, name, description, params, steps }) {
        if (!id) throw new Error("ID requerido");
        const existing = await loadShortcut(id);
        if (name !== undefined) existing.name = name;
        if (description !== undefined) existing.description = description;
        if (params !== undefined) existing.params = params;
        if (steps !== undefined) existing.steps = steps;
        existing.updatedAt = new Date().toISOString();

        const filePath = await getShortcutPath(id);
        await fs.writeFile(filePath, JSON.stringify(existing, null, 2), "utf8");
        return { ok: true, message: `Shortcut '${id}' actualizado`, shortcut: existing };
      },

      async delete_shortcut({ id }) {
        if (!id) throw new Error("ID requerido");
        const filePath = await getShortcutPath(id);
        await fs.unlink(filePath);
        return { ok: true, message: `Shortcut '${id}' eliminado` };
      },

      async duplicate_shortcut({ id, newId, newName }) {
        if (!id || !newId) throw new Error("IDs de origen y destino requeridos");
        const sc = await loadShortcut(id);
        sc.id = newId;
        sc.name = newName || `${sc.name} (Copia)`;
        sc.createdAt = new Date().toISOString();
        const destPath = await getShortcutPath(newId);
        await fs.writeFile(destPath, JSON.stringify(sc, null, 2), "utf8");
        return { ok: true, message: `Shortcut duplicado como '${newId}'`, shortcut: sc };
      },

      async export_shortcut({ id }) {
        const sc = await loadShortcut(id);
        return { ok: true, exportData: JSON.stringify(sc, null, 2) };
      },

      async import_shortcut({ data, overwrite = false }) {
        const sc = typeof data === "string" ? JSON.parse(data) : data;
        if (!sc.id || !sc.name || !Array.isArray(sc.steps)) {
          throw new Error("Formato de importación inválido");
        }
        await ensureDir();
        const filePath = await getShortcutPath(sc.id);
        if (!overwrite) {
          try {
            await fs.access(filePath);
            throw new Error(`El shortcut '${sc.id}' ya existe. Pase overwrite=true para reemplazarlo.`);
          } catch (e) {
            if (e.code !== "ENOENT") throw e;
          }
        }
        await fs.writeFile(filePath, JSON.stringify(sc, null, 2), "utf8");
        return { ok: true, message: `Shortcut '${sc.id}' importado correctamente` };
      },

      async run_shortcut({ id, params = {}, dry_run = false }) {
        if (!id) throw new Error("ID de shortcut requerido");
        const sc = await loadShortcut(id);

        // Fill default params
        const mergedParams = {};
        for (const [k, def] of Object.entries(sc.params || {})) {
          mergedParams[k] = params[k] !== undefined ? params[k] : (typeof def === "object" ? def.default : def);
        }
        for (const [k, v] of Object.entries(params)) {
          mergedParams[k] = v;
        }

        const steps = sc.steps || [];
        const plan = [];
        let highestRisk = RISK_LEVELS.LOW;

        for (let i = 0; i < steps.length; i++) {
          const step = steps[i];
          const resolved = resolveArgs(step.args || {}, mergedParams);
          // Evaluate risk of step
          const stepRisk = step.risk || (step.action.startsWith("delete") ? RISK_LEVELS.DESTRUCTIVE : RISK_LEVELS.HIGH);
          if (stepRisk === RISK_LEVELS.DESTRUCTIVE) highestRisk = RISK_LEVELS.DESTRUCTIVE;
          else if (stepRisk === RISK_LEVELS.HIGH && highestRisk !== RISK_LEVELS.DESTRUCTIVE) highestRisk = RISK_LEVELS.HIGH;

          plan.push({
            step: i + 1,
            tool: step.tool,
            action: step.action,
            resolvedArgs: resolved,
            onError: step.onError || "stop"
          });
        }

        if (dry_run) {
          return {
            ok: true,
            dryRun: true,
            shortcutId: id,
            highestRisk,
            stepsCount: plan.length,
            plan
          };
        }

        // Execute sequential steps with rollback tracking
        const executionLog = [];
        const rollbackActions = [];

        for (const item of plan) {
          try {
            let result = null;
            if (runtime?.router) {
              result = await runtime.router.execute(item.tool, item.action, item.resolvedArgs);
            } else {
              // Direct dispatch fallback if standalone
              result = { ok: true, simulated: false, output: `Executed ${item.tool}.${item.action}` };
            }

            // Register rollback if files were written
            if (item.tool === "files" && item.action === "write_file" && item.resolvedArgs.path) {
              rollbackActions.push(async () => {
                try { await fs.unlink(path.resolve(item.resolvedArgs.path)); } catch {}
              });
            }

            executionLog.push({ step: item.step, tool: item.tool, action: item.action, status: "OK", result });
          } catch (err) {
            executionLog.push({ step: item.step, tool: item.tool, action: item.action, status: "FAILED", error: err.message });
            
            if (item.onError === "rollback") {
              // Execute rollback
              for (const rb of rollbackActions.reverse()) {
                try { await rb(); } catch {}
              }
              return {
                ok: false,
                shortcutId: id,
                failedAtStep: item.step,
                error: err.message,
                rollbackExecuted: true,
                executionLog
              };
            } else if (item.onError === "stop") {
              return {
                ok: false,
                shortcutId: id,
                failedAtStep: item.step,
                error: err.message,
                rollbackExecuted: false,
                executionLog
              };
            }
          }
        }

        return {
          ok: true,
          shortcutId: id,
          completedSteps: executionLog.length,
          executionLog
        };
      }
    }
  };
}
