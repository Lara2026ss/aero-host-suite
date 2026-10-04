/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * tools/guide.mjs — Lazy Catalog and Capability Discovery Engine
 */
import { getVersion, PRODUCT_NAME } from "../core/version.mjs";
import { RISK_LEVELS } from "../core/permissions.mjs";

export function createGuideDomain({ runtime }) {
  return {
    name: "guide",
    description: "Catálogo perezoso de herramientas, documentación compacta y estado general de Fluxer XZ.",
    permissions: {
      get_catalog: RISK_LEVELS.LOW,
      get_tool_guide: RISK_LEVELS.LOW,
      get_system_status: RISK_LEVELS.LOW
    },
    actions: {
      async get_catalog({ filter = "" }) {
        const reg = runtime?.registry;
        const catalog = [];

        if (reg) {
          for (const [name, mod] of reg.modules.entries()) {
            if (filter && !name.includes(filter) && !mod.description.includes(filter)) continue;
            const actionNames = Object.keys(mod.actions || {}).sort();
            catalog.push({
              tool: name,
              description: mod.description,
              actionsCount: actionNames.length,
              actions: actionNames.map((act) => ({
                action: act,
                risk: mod.permissions?.[act] || "low"
              }))
            });
          }
        }

        return {
          ok: true,
          product: PRODUCT_NAME,
          version: getVersion(),
          toolsCount: catalog.length,
          catalog
        };
      },

      async get_tool_guide({ tool }) {
        if (!tool) throw new Error("Parámetro 'tool' requerido.");
        const reg = runtime?.registry;
        const mod = reg?.modules?.get(tool);
        if (!mod) throw new Error(`Herramienta '${tool}' no encontrada en el registro.`);

        const actionsDetail = {};
        for (const [actName, fn] of Object.entries(mod.actions || {})) {
          const sig = mod.actionSignatures?.[actName] || "{}";
          actionsDetail[actName] = {
            signature: sig,
            risk: mod.permissions?.[actName] || "low"
          };
        }

        return {
          ok: true,
          tool,
          description: mod.description,
          actions: actionsDetail
        };
      },

      async get_system_status() {
        const reg = runtime?.registry;
        return {
          ok: true,
          product: PRODUCT_NAME,
          version: getVersion(),
          toolsLoaded: reg ? reg.modules.size : 0,
          compactMode: runtime?.compact?.getMode() || "auto",
          isElevated: runtime?.security?.isElevated || false,
          status: "healthy"
        };
      }
    }
  };
}
