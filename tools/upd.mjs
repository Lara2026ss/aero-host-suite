/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * tools/upd.mjs — Official Tool Interface for Autonomous Remote Updates
 */
import { UpdaterEngine } from "../updater/updater.mjs";
import { RISK_LEVELS } from "../core/permissions.mjs";

export function createUpdDomain({ runtime }) {
  const engine = new UpdaterEngine({ rootDir: runtime?.rootDir || process.cwd() });

  return {
    name: "upd",
    description: "Centro oficial de actualización remota autónoma, verificación SHA-256 y rollback atómico.",
    permissions: {
      check_update: RISK_LEVELS.LOW,
      apply_update: RISK_LEVELS.HIGH,
      rollback_update: RISK_LEVELS.HIGH,
      get_status: RISK_LEVELS.LOW
    },
    actions: {
      async check_update({ endpoint = null }) {
        return engine.check(endpoint);
      },

      async apply_update({ force = false, expectedSha256 = null, packageUrl = null }) {
        return engine.apply({ force, expectedSha256, packageUrl });
      },

      async rollback_update({ backupPath = null }) {
        return engine.rollback({ backupPath });
      },

      async get_status() {
        return engine.getStatus();
      }
    }
  };
}
