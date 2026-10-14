/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
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
      run_command: RISK_LEVELS.LOW,
      run_background: RISK_LEVELS.LOW,
      list_jobs: RISK_LEVELS.LOW,
      kill_job: RISK_LEVELS.LOW,
      run_script: RISK_LEVELS.LOW
    },
    actions: {
      async run_command(args = {}) {
        const command = args.command || args.cmd || args.script || args.code;
        const cwd = args.cwd || process.cwd();
        const timeoutMs = args.timeoutMs ? Number(args.timeoutMs) : 60000;
        const maxOutputLines = args.maxOutputLines ? Number(args.maxOutputLines) : 120;

        if (!command) throw new Error("Parámetro 'command' requerido.");

        return new Promise((resolve) => {
          const shell = process.platform === "win32" ? "powershell.exe" : "/bin/bash";
          const shellArgs = process.platform === "win32" ? ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", command] : ["-c", command];

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

      async run_background(args = {}) {
        const command = args.command || args.cmd || args.script || args.code;
        const cwd = args.cwd || process.cwd();
        if (!command) throw new Error("Parámetro 'command' requerido.");
        jobCounter++;
        const jobId = `job#${jobCounter}`;
        const shell = process.platform === "win32" ? "powershell.exe" : "/bin/bash";
        const shellArgs = process.platform === "win32" ? ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", command] : ["-c", command];

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

      async kill_job(args = {}) {
        const jobId = args.jobId || args.id;
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

      async run_script(args = {}) {
        const script = args.script || args.code || args.content;
        const language = args.language || args.lang || "powershell";
        const scriptArgs = args.args || [];
        if (!script) throw new Error("Contenido de 'script' requerido.");
        const ext = language === "cmd" ? ".bat" : ".ps1";
        const tempPath = path.join(os.tmpdir(), `aero_host_script_${Date.now()}${ext}`);
        await fs.writeFile(tempPath, script, "utf8");

        try {
          let command = "";
          if (language === "cmd") {
            command = `cmd.exe /c "${tempPath}" ${scriptArgs.join(" ")}`;
          } else {
            command = `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "${tempPath}" ${scriptArgs.join(" ")}`;
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
