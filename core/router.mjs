/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * core/router.mjs — Secure Tool Execution Router
 */
import { security } from "./permissions.mjs";
import { compact } from "./compact.mjs";

export class Router {
  constructor(runtime) {
    this.runtime = runtime;
    if (this.runtime) {
      this.runtime.router = this;
    }
  }

  async execute(tool, action, params = {}) {
    const reg = this.runtime?.registry;
    if (!reg) throw new Error("Registry no inicializado en el router.");

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

    // Determine target resource or context for permission check
    const resource = params.path || params.file || params.targetPath || params.src || params.command || params.query || "";
    const reason = params.reason || `Invocación de acción ${tool}.${action}`;

    // Security Gate Enforcement
    const permResult = await security.checkPermission({
      tool,
      action,
      risk: resolved.risk,
      resource,
      reason
    });

    if (!permResult.allowed) {
      return {
        ok: false,
        error: permResult.reason,
        securityBlocked: true,
        risk: resolved.risk
      };
    }

    // Resolve any stash references in string parameters
    const sanitizedParams = { ...params };
    for (const [k, v] of Object.entries(sanitizedParams)) {
      if (typeof v === "string" && v.startsWith("stash#")) {
        sanitizedParams[k] = compact.resolveContent(v);
      }
    }

    // Handler Execution
    try {
      const result = await resolved.handler.call(resolved.domain, sanitizedParams);
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
