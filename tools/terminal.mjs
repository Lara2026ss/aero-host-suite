/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * tools/terminal.mjs — High-Performance Process and Terminal Execution Engine
 */
import { exec, spawn } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { compact } from "../core/compact.mjs";
import { RISK_LEVELS } from "../core/permissions.mjs";

export function createTerminalDomain({ runtime }) {
  const jobs = new Map(); // jobId -> { child, command, startedAt, stdout, stderr, status }
  let jobCounter = 0;

  return {
    name: "terminal",
    description: "Ejecución de comandos, procesos en segundo plano y scripts de sistema.",
    permissions: {
      run_command: RISK_LEVELS.HIGH,
      run_background: RISK_LEVELS.HIGH,
      list_jobs: RISK_LEVELS.LOW,
      kill_job: RISK_LEVELS.HIGH,
      run_script: RISK_LEVELS.HIGH
    },
    actions: {
      async run_command({ command, cwd = process.cwd(), timeoutMs = 30000, maxOutputLines = 80 }) {
        if (!command) throw new Error("Parámetro 'command' requerido.");

        return new Promise((resolve) => {
          const shell = process.platform === "win32" ? "powershell.exe" : "/bin/bash";
          const shellArgs = process.platform === "win32" ? ["-NoProfile", "-Command", command] : ["-c", command];

          const start = Date.now();
          const child = spawn(shell, shellArgs, { cwd, windowsHide: true });
          let stdout = "";
          let stderr = "";

          const timer = setTimeout(() => {
            child.kill("SIGKILL");
            resolve({
              ok: false,
              timedOut: true,
              exitCode: -1,
              durationMs: Date.now() - start,
              error: `El comando excedió el límite de tiempo de ${timeoutMs}ms.`
            });
          }, timeoutMs);

          child.stdout.on("data", (d) => { stdout += d.toString(); });
          child.stderr.on("data", (d) => { stderr += d.toString(); });

          child.on("close", (code) => {
            clearTimeout(timer);
            const durationMs = Date.now() - start;
            resolve({
              ok: code === 0,
              exitCode: code,
              durationMs,
              stdout: compact.truncateOutput(stdout.trim(), maxOutputLines),
              stderr: compact.truncateOutput(stderr.trim(), 40)
            });
          });

          child.on("error", (err) => {
            clearTimeout(timer);
            resolve({
              ok: false,
              exitCode: -1,
              durationMs: Date.now() - start,
              error: err.message
            });
          });
        });
      },

      async run_background({ command, cwd = process.cwd() }) {
        if (!command) throw new Error("Parámetro 'command' requerido.");
        jobCounter++;
        const jobId = `job#${jobCounter}`;
        const shell = process.platform === "win32" ? "powershell.exe" : "/bin/bash";
        const shellArgs = process.platform === "win32" ? ["-NoProfile", "-Command", command] : ["-c", command];

        const child = spawn(shell, shellArgs, { cwd, windowsHide: true, detached: true });
        const jobInfo = {
          id: jobId,
          command,
          pid: child.pid,
          startedAt: new Date().toISOString(),
          status: "running",
          stdout: "",
          stderr: ""
        };

        child.stdout.on("data", (d) => { jobInfo.stdout += d.toString(); });
        child.stderr.on("data", (d) => { jobInfo.stderr += d.toString(); });
        child.on("close", (code) => {
          jobInfo.status = code === 0 ? "completed" : "failed";
          jobInfo.exitCode = code;
          jobInfo.completedAt = new Date().toISOString();
        });

        jobs.set(jobId, { child, info: jobInfo });
        return { ok: true, jobId, pid: child.pid, status: "running", command };
      },

      async list_jobs() {
        const list = [];
        for (const [id, item] of jobs.entries()) {
          list.push({
            id,
            pid: item.info.pid,
            command: item.info.command,
            status: item.info.status,
            startedAt: item.info.startedAt,
            completedAt: item.info.completedAt,
            exitCode: item.info.exitCode
          });
        }
        return { ok: true, count: list.length, jobs: list };
      },

      async kill_job({ jobId }) {
        if (!jobId) throw new Error("Parámetro 'jobId' requerido.");
        const job = jobs.get(jobId);
        if (!job) throw new Error(`Trabajo '${jobId}' no encontrado.`);

        try {
          if (job.child && !job.child.killed) {
            job.child.kill("SIGKILL");
          }
          job.info.status = "killed";
          return { ok: true, message: `Trabajo '${jobId}' detenido.` };
        } catch (e) {
          throw new Error(`Error al detener trabajo: ${e.message}`);
        }
      },

      async run_script({ script, language = "powershell", args = [] }) {
        if (!script) throw new Error("Contenido de 'script' requerido.");
        const ext = language === "cmd" ? ".bat" : ".ps1";
        const tempPath = path.join(os.tmpdir(), `fluxer_script_${Date.now()}${ext}`);
        await fs.writeFile(tempPath, script, "utf8");

        try {
          let command = "";
          if (language === "cmd") {
            command = `cmd.exe /c "${tempPath}" ${args.join(" ")}`;
          } else {
            command = `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "${tempPath}" ${args.join(" ")}`;
          }

          const res = await this.run_command({ command });
          return res;
        } finally {
          await fs.unlink(tempPath).catch(() => {});
        }
      }
    }
  };
}
