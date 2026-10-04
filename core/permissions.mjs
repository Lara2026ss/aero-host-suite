/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * core/permissions.mjs — Unified Security & Elevation Router Engine
 */
import fs from "node:fs/promises";
import path from "node:path";
import { promptSecurityConfirmation, sendNativeNotification, SECURITY_TITLE } from "./notify.mjs";

export const RISK_LEVELS = Object.freeze({
  LOW: "low",
  MEDIUM: "medium",
  HIGH: "high",
  DESTRUCTIVE: "destructive"
});

export class SecurityManager {
  constructor(options = {}) {
    this.rootDir = options.rootDir || process.cwd();
    this.auditLogPath = path.join(this.rootDir, "data", "security", "audit.log");
    this.sessionPermissions = new Set(); // Stores "tool:action" or "tool:*"
    this.isElevated = false;
    this.elevationScope = null;
    this.blockedPaths = [
      /[\\\/]windows[\\\/]system32/i,
      /[\\\/]windows[\\\/]syswow64/i,
      /[\\\/]config[\\\/](sam|system|security|software|default)$/i,
      /[\\\/]\.ssh[\\\/](id_rsa|id_ecdsa|id_ed25519|authorized_keys)$/i,
      /[\\\/]\.aws[\\\/]credentials$/i,
      /[\\\/]\.azure[\\\/]/i,
      /[\\\/]\.kube[\\\/]config$/i,
      /[\\\/]drivers[\\\/]etc[\\\/]hosts$/i,
      /[\\\/]etc[\\\/](shadow|passwd|sudoers)$/i
    ];
  }

  setElevated(value, scope = "*") {
    this.isElevated = Boolean(value);
    this.elevationScope = this.isElevated ? scope : null;
    this.logAudit({
      event: "elevation_change",
      isElevated: this.isElevated,
      scope: this.elevationScope,
      timestamp: new Date().toISOString()
    });
  }

  revokeElevation() {
    this.setElevated(false, null);
    this.sessionPermissions.clear();
    this.logAudit({
      event: "elevation_revoked",
      timestamp: new Date().toISOString()
    });
  }

  getSecurityStatus() {
    return {
      isElevated: this.isElevated,
      elevationScope: this.elevationScope,
      sessionPermissionsCount: this.sessionPermissions.size,
      activePermissions: Array.from(this.sessionPermissions),
      blockedPathsCount: this.blockedPaths.length
    };
  }

  isPathBlocked(targetPath) {
    if (!targetPath) return false;
    const normalized = path.resolve(String(targetPath));
    for (const pattern of this.blockedPaths) {
      if (pattern.test(normalized)) return true;
    }
    return false;
  }

  async checkPermission({ tool, action, risk = RISK_LEVELS.LOW, resource = "", reason = "" }) {
    const key = `${tool}:${action}`;
    const toolWildcard = `${tool}:*`;

    // 1. Strict System Path Block
    if (resource && this.isPathBlocked(resource)) {
      await this.logAudit({
        tool, action, risk, resource, decision: "blocked_system_integrity",
        timestamp: new Date().toISOString()
      });
      return {
        allowed: false,
        reason: `Acceso denegado: El recurso '${resource}' está protegido por la integridad del sistema operativo.`
      };
    }

    // 2. Low Risk Actions: Always permitted
    if (risk === RISK_LEVELS.LOW) {
      await this.logAudit({
        tool, action, risk, resource, decision: "allowed_low_risk",
        timestamp: new Date().toISOString()
      });
      return { allowed: true, risk };
    }

    // 3. Check Session Permissions
    if (this.sessionPermissions.has(key) || this.sessionPermissions.has(toolWildcard)) {
      if (risk !== RISK_LEVELS.DESTRUCTIVE) {
        await this.logAudit({
          tool, action, risk, resource, decision: "allowed_session_cache",
          timestamp: new Date().toISOString()
        });
        return { allowed: true, risk };
      }
    }

    // 4. Elevation Check: If session is elevated, MEDIUM and HIGH proceed
    if (this.isElevated) {
      if (this.elevationScope === "*" || this.elevationScope === tool) {
        if (risk !== RISK_LEVELS.DESTRUCTIVE) {
          await this.logAudit({
            tool, action, risk, resource, decision: "allowed_elevated_session",
            timestamp: new Date().toISOString()
          });
          return { allowed: true, risk };
        }
      }
    }

    // 5. Explicit Confirmation Required (HIGH or DESTRUCTIVE or un-elevated MEDIUM)
    // Send high-priority notification toast
    if (risk === RISK_LEVELS.HIGH || risk === RISK_LEVELS.DESTRUCTIVE) {
      sendNativeNotification(
        SECURITY_TITLE,
        `Acción de nivel ${risk.toUpperCase()}: ${tool}.${action} requiere confirmación.`
      ).catch(() => {});
    }

    const decision = await promptSecurityConfirmation({
      tool,
      action,
      resource,
      reason,
      risk
    });

    if (decision === "allow_session") {
      this.sessionPermissions.add(key);
      await this.logAudit({
        tool, action, risk, resource, decision: "granted_session",
        timestamp: new Date().toISOString()
      });
      return { allowed: true, risk };
    }

    if (decision === "allow_once") {
      await this.logAudit({
        tool, action, risk, resource, decision: "granted_once",
        timestamp: new Date().toISOString()
      });
      return { allowed: true, risk };
    }

    // Default: Denied
    await this.logAudit({
      tool, action, risk, resource, decision: "denied_by_user",
      timestamp: new Date().toISOString()
    });
    return {
      allowed: false,
      reason: `Acción cancelada: El usuario rechazó el permiso para ejecutar '${tool}.${action}'.`
    };
  }

  async logAudit(entry) {
    try {
      const logDir = path.dirname(this.auditLogPath);
      await fs.mkdir(logDir, { recursive: true });
      const line = JSON.stringify(entry) + "\n";
      await fs.appendFile(this.auditLogPath, line, "utf8");
    } catch {}
  }
}

export const security = new SecurityManager();
