/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * tools/security.mjs — Security, Elevation & Cryptographic Audit Tool
 */
import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { security, RISK_LEVELS } from "../core/permissions.mjs";

export function createSecurityDomain({ runtime }) {
  return {
    name: "security",
    description: "Gestión de elevación de privilegios, estado de seguridad, hashes y auditoría de accesos.",
    permissions: {
      get_status: RISK_LEVELS.LOW,
      request_elevation: RISK_LEVELS.HIGH,
      revoke_elevation: RISK_LEVELS.LOW,
      audit_log: RISK_LEVELS.LOW,
      hash_data: RISK_LEVELS.LOW,
      verify_hash: RISK_LEVELS.LOW
    },
    actions: {
      async get_status() {
        return {
          ok: true,
          product: "Fluxer XZ · Seguridad",
          isElevated: security.isElevated,
          elevationScope: security.elevationScope,
          sessionGrantsCount: security.sessionPermissions.size,
          blockedRulesCount: security.blockedPaths.length,
          auditLogPath: security.auditLogPath
        };
      },

      async request_elevation({ scope = "*" }) {
        security.setElevated(true, scope);
        return {
          ok: true,
          message: `Elevación concedida para el ámbito: ${scope}. Las acciones de nivel ALTO procederán sin interrupciones durante esta sesión.`,
          scope,
          isElevated: true
        };
      },

      async revoke_elevation() {
        security.setElevated(false, null);
        security.sessionPermissions.clear();
        return { ok: true, message: "Elevación revocada y permisos de sesión reseteados.", isElevated: false };
      },

      async audit_log({ limit = 50 }) {
        try {
          const raw = await fs.readFile(security.auditLogPath, "utf8");
          const lines = raw.trim().split("\n").filter(Boolean);
          const entries = lines.slice(-limit).map((l) => {
            try { return JSON.parse(l); } catch { return { raw: l }; }
          });
          return { ok: true, count: entries.length, log: entries.reverse() };
        } catch (e) {
          if (e.code === "ENOENT") {
            return { ok: true, count: 0, log: [] };
          }
          throw e;
        }
      },

      async hash_data({ data, filePath = null, algorithm = "sha256" }) {
        let content = data;
        if (filePath) {
          content = await fs.readFile(path.resolve(filePath));
        }
        if (content === undefined || content === null) {
          throw new Error("Especifique 'data' o 'filePath'.");
        }
        const hash = crypto.createHash(algorithm).update(content).digest("hex");
        return { ok: true, algorithm, hash, bytes: Buffer.byteLength(content) };
      },

      async verify_hash({ expectedHash, data, filePath = null, algorithm = "sha256" }) {
        if (!expectedHash) throw new Error("Parámetro 'expectedHash' requerido.");
        const res = await this.hash_data({ data, filePath, algorithm });
        const matches = res.hash.toLowerCase() === expectedHash.trim().toLowerCase();
        return { ok: true, matches, calculatedHash: res.hash, expectedHash };
      }
    }
  };
}
