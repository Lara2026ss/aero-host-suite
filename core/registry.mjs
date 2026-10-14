/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
 * core/registry.mjs — Slim & High-Performance Tool Registry
 */
import path from "node:path";
import fs from "node:fs/promises";
import crypto from "node:crypto";
import { pathToFileURL } from "node:url";

function extractParamHint(fn) {
  if (typeof fn !== "function") return "{}";
  const str = fn.toString();
  const match = str.match(/(?:async\s*)?(?:function\s*)?(?:[^(]*)\(\s*\{([^}]*)\}/);
  if (!match) return "{}";
  const params = match[1]
    .split(",")
    .map((item) => {
      const noDefault = item.split("=")[0].trim();
      const propKey = noDefault.split(":")[0].trim();
      return propKey.replace(/[^a-zA-Z0-9_]/g, "");
    })
    .filter(Boolean);
  return params.length ? `{ ${params.join(", ")} }` : "{}";
}

export class Registry {
  constructor(runtime) {
    this.runtime = runtime;
    this.modules = new Map();
    if (this.runtime) {
      this.runtime.registry = this;
    }
  }

  async load() {
    const runtime = this.runtime;
    const toolsDir = path.resolve(runtime?.rootDir || process.cwd(), "tools");

    const toolFactories = [
      { file: "files.mjs", factory: "createFilesDomain" },
      { file: "terminal.mjs", factory: "createTerminalDomain" },
      { file: "system.mjs", factory: "createSystemDomain" },
      { file: "web.mjs", factory: "createWebDomain" },
      { file: "documents.mjs", factory: "createDocumentsDomain" },
      { file: "media.mjs", factory: "createMediaDomain" },
      { file: "shortcuts.mjs", factory: "createShortcutsDomain" },
      { file: "upd.mjs", factory: "createUpdDomain" },
      { file: "feedback.mjs", factory: "createFeedbackDomain" },
      { file: "guide.mjs", factory: "createGuideDomain" }
    ];

    const context = { runtime, path, fs, crypto, router: runtime?.router };

    for (const item of toolFactories) {
      const fullPath = path.join(toolsDir, item.file);
      try {
        const fileUrl = pathToFileURL(fullPath).href;
        const mod = await import(`${fileUrl}?t=${Date.now()}`);
        const fn = mod[item.factory] || Object.values(mod).find((exp) => typeof exp === "function");
        if (typeof fn === "function") {
          const domainObj = fn(context);
          if (domainObj && domainObj.name && domainObj.actions) {
            domainObj.actionSignatures = {};
            for (const [k, actionFn] of Object.entries(domainObj.actions)) {
              domainObj.actionSignatures[k] = extractParamHint(actionFn);
            }
            this.modules.set(domainObj.name, domainObj);
          }
        }
      } catch (err) {
        console.error(`[Registry] Error cargando herramienta '${item.file}':`, err.message);
      }
    }

    return this.modules.size;
  }

  moduleNames() {
    return [...this.modules.keys()];
  }

  get(name) {
    return this.modules.get(name) || null;
  }

  actionsFor(name) {
    const mod = this.modules.get(name);
    return mod ? Object.keys(mod.actions || {}).sort() : [];
  }

  resolve(tool, action) {
    const mod = this.modules.get(tool);
    if (!mod) return null;
    const fn = mod.actions?.[action];
    if (!fn) return null;
    return {
      domain: mod,
      handler: fn,
      risk: mod.permissions?.[action] || "low",
      signature: mod.actionSignatures?.[action] || "{}"
    };
  }
}
