/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
 * core/router.mjs — Unrestricted Tool Execution & Input/Output Normalization Engine
 */
import { compact } from "./compact.mjs";

export class Router {
  constructor(runtime) {
    this.runtime = runtime;
    if (this.runtime) {
      this.runtime.router = this;
    }
  }

  normalizeInputs(tool, action, params = {}) {
    const p = { ...params };

    // Common aliases across all tools
    if (!p.path && (p.file || p.filePath || p.targetPath || p.destination || p.doc)) {
      p.path = p.file || p.filePath || p.targetPath || p.destination || p.doc;
    }
    if (p.content !== undefined && p.text === undefined && (tool === "documents" || tool === "files")) {
      p.text = p.content;
    }
    if (p.command === undefined && (p.cmd || p.script || p.code)) {
      p.command = p.cmd || p.script || p.code;
    }

    // Specific domain normalization
    if (tool === "guide") {
      if (!p.tool && (p.name || p.query || p.domain || p.target)) {
        p.tool = p.name || p.query || p.domain || p.target;
      }
      if (!p.query && (p.tool || p.name || p.filter)) {
        p.query = p.tool || p.name || p.filter;
      }
    }

    if (tool === "documents") {
      if (!p.app && (p.targetApp || p.application)) {
        p.app = p.targetApp || p.application;
      }
    }

    return p;
  }

  async execute(tool, action, params = {}) {
    const reg = this.runtime?.registry;
    if (!reg) throw new Error("Registry no inicializado en el router.");

    // Normalize inputs before resolution and dispatch
    const normalizedParams = this.normalizeInputs(tool, action, params);

    const resolved = reg.resolve(tool, action);
    if (!resolved) {
      const availableTools = reg.moduleNames();
      const mod = reg.get(tool);
      if (!mod) {
        throw new Error(`Herramienta '${tool}' no encontrada. Disponibles: ${availableTools.join(", ")}`);
      }
      const availableActions = Object.keys(mod.actions || {});
      throw new Error(`Acción '${action}' no encontrada en herramienta '${tool}'. Disponibles: ${availableActions.join(", ")}`);
    }

    // Resolve any stash references in string parameters
    const sanitizedParams = { ...normalizedParams };
    for (const [k, v] of Object.entries(sanitizedParams)) {
      if (typeof v === "string" && v.startsWith("stash#")) {
        sanitizedParams[k] = compact.resolveContent(v);
      }
    }

    // Handler Execution - Unrestricted, zero security blocks
    try {
      const result = await resolved.handler.call(resolved.domain, sanitizedParams);
      if (result && typeof result === "object" && result.ok === undefined) {
        result.ok = true;
      }
      return result;
    } catch (err) {
      return {
        ok: false,
        error: err.message,
        code: err.code || "TOOL_EXECUTION_ERROR"
      };
    }
  }
}
