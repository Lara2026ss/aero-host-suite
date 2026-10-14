/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
 * tools/system.mjs — Comprehensive System, Diagnostics, Network & Package Management
 */
import os from "node:os";
import dns from "node:dns/promises";
import net from "node:net";
import { exec } from "node:child_process";
import { promisify } from "node:util";
import { RISK_LEVELS } from "../core/permissions.mjs";

const execAsync = promisify(exec);

export function createSystemDomain({ runtime }) {
  return {
    name: "system",
    description: "Estado del sistema, procesos, servicios, diagnóstico de red y gestor de paquetes.",
    permissions: {
      get_info: RISK_LEVELS.LOW,
      list_processes: RISK_LEVELS.LOW,
      kill_process: RISK_LEVELS.HIGH,
      manage_service: RISK_LEVELS.HIGH,
      diagnose_network: RISK_LEVELS.LOW,
      list_packages: RISK_LEVELS.LOW,
      install_package: RISK_LEVELS.HIGH,
      remove_package: RISK_LEVELS.HIGH
    },
    actions: {
      async get_info({ includeEnv = false }) {
        const cpus = os.cpus();
        const totalMem = os.totalmem();
        const freeMem = os.freemem();
        const usedMem = totalMem - freeMem;

        const info = {
          hostname: os.hostname(),
          platform: os.platform(),
          release: os.release(),
          arch: os.arch(),
          uptimeHours: (os.uptime() / 3600).toFixed(2),
          cpu: {
            model: cpus[0]?.model || "Unknown",
            cores: cpus.length,
            speedMhz: cpus[0]?.speed || 0
          },
          memory: {
            totalMb: Math.round(totalMem / (1024 * 1024)),
            usedMb: Math.round(usedMem / (1024 * 1024)),
            freeMb: Math.round(freeMem / (1024 * 1024)),
            percentUsed: `${Math.round((usedMem / totalMem) * 100)}%`
          },
          networkInterfaces: Object.keys(os.networkInterfaces())
        };

        if (includeEnv) {
          // Filter out sensitive tokens
          const safeEnv = {};
          for (const [k, v] of Object.entries(process.env)) {
            if (/token|secret|password|key|auth/i.test(k)) {
              safeEnv[k] = "[PROTECTED]";
            } else {
              safeEnv[k] = v;
            }
          }
          info.env = safeEnv;
        }

        return { ok: true, system: info };
      },

      async list_processes({ limit = 30, filter = "" }) {
        let cmd = "Get-Process | Select-Object -First 100 Id, ProcessName, WorkingSet64, CPU | ConvertTo-Json";
        if (filter) {
          cmd = `Get-Process -Name "*${filter}*" -ErrorAction SilentlyContinue | Select-Object -First 50 Id, ProcessName, WorkingSet64, CPU | ConvertTo-Json`;
        }

        const { stdout } = await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
        let data = [];
        try {
          const parsed = JSON.parse(stdout);
          data = Array.isArray(parsed) ? parsed : [parsed];
        } catch {
          data = [];
        }

        const formatted = data.slice(0, limit).map((p) => ({
          pid: p.Id,
          name: p.ProcessName,
          memMb: Math.round((p.WorkingSet64 || 0) / (1024 * 1024)),
          cpu: p.CPU ? Number(p.CPU).toFixed(1) : "0.0"
        }));

        return { ok: true, count: formatted.length, processes: formatted };
      },

      async kill_process({ pid, name, force = true }) {
        if (!pid && !name) throw new Error("Especifique 'pid' o 'name' para terminar el proceso.");
        let cmd = "";
        if (pid) {
          const cleanPid = parseInt(pid, 10);
          if (isNaN(cleanPid)) throw new Error("El 'pid' debe ser numérico.");
          cmd = `Stop-Process -Id ${cleanPid} ${force ? "-Force" : ""}`;
        } else {
          const safeName = String(name).replace(/[`"$;&|<>]/g, "");
          cmd = `Stop-Process -Name "${safeName}" ${force ? "-Force" : ""}`;
        }

        await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
        return { ok: true, message: `Proceso ${pid || name} terminado.` };
      },

      async manage_service({ serviceName, action = "status" }) {
        if (!serviceName) throw new Error("Parámetro 'serviceName' requerido.");
        const validActions = ["status", "start", "stop", "restart"];
        if (!validActions.includes(action)) throw new Error(`Acción inválida. Opciones: ${validActions.join(", ")}`);
        const safeService = String(serviceName).replace(/[`"$;&|<>]/g, "");

        let cmd = "";
        if (action === "status") {
          cmd = `Get-Service -Name "${safeService}" | Select-Object Name, DisplayName, Status | ConvertTo-Json`;
          const { stdout } = await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
          return { ok: true, service: JSON.parse(stdout) };
        } else if (action === "start") {
          cmd = `Start-Service -Name "${safeService}"`;
        } else if (action === "stop") {
          cmd = `Stop-Service -Name "${safeService}" -Force`;
        } else if (action === "restart") {
          cmd = `Restart-Service -Name "${safeService}" -Force`;
        }

        await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
        return { ok: true, message: `Servicio '${safeService}' ejecutó acción '${action}'.` };
      },

      async diagnose_network({ host = "1.1.1.1", port = null }) {
        const results = { host };

        // DNS Lookup
        try {
          const addresses = await dns.lookup(host);
          results.dns = { resolvedIp: addresses.address, family: addresses.family };
        } catch (e) {
          results.dns = { error: e.message };
        }

        // Port connectivity check
        if (port) {
          const portNum = parseInt(port, 10);
          results.portCheck = await new Promise((resolve) => {
            const socket = new net.Socket();
            socket.setTimeout(3000);
            socket.on("connect", () => {
              socket.destroy();
              resolve({ port: portNum, open: true });
            });
            socket.on("timeout", () => {
              socket.destroy();
              resolve({ port: portNum, open: false, reason: "timeout" });
            });
            socket.on("error", (err) => {
              resolve({ port: portNum, open: false, reason: err.message });
            });
            socket.connect(portNum, host);
          });
        }

        // Ping test
        try {
          const { stdout } = await execAsync(`ping -n 2 -w 1000 ${host}`);
          results.ping = { ok: !stdout.includes("100% loss"), summary: stdout.split(/\r?\n/).slice(-3).join(" ") };
        } catch {
          results.ping = { ok: false };
        }

        return { ok: true, network: results };
      },

      async list_packages({ manager = "winget", limit = 30 }) {
        let cmd = manager === "npm" ? "npm list -g --depth=0 --json" : "winget list -n 30";
        try {
          const { stdout } = await execAsync(cmd);
          return { ok: true, manager, output: stdout.slice(0, 4000) };
        } catch (e) {
          return { ok: false, error: e.message };
        }
      },

      async install_package({ name, manager = "winget" }) {
        if (!name) throw new Error("Parámetro 'name' de paquete requerido.");
        const cmd = manager === "npm"
          ? `npm install -g ${name}`
          : `winget install --id "${name}" --accept-source-agreements --accept-package-agreements -e`;

        const { stdout, stderr } = await execAsync(cmd);
        return { ok: true, message: `Paquete '${name}' instalado exitosamente.`, output: stdout || stderr };
      },

      async remove_package({ name, manager = "winget" }) {
        if (!name) throw new Error("Parámetro 'name' de paquete requerido.");
        const cmd = manager === "npm" ? `npm uninstall -g ${name}` : `winget uninstall --id "${name}" -e`;
        const { stdout, stderr } = await execAsync(cmd);
        return { ok: true, message: `Paquete '${name}' desinstalado.`, output: stdout || stderr };
      }
    }
  };
}
