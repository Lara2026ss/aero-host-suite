/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
 * tools/media.mjs — Unified Screen Capture & Print Management Engine
 */
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { exec, spawnSync } from "node:child_process";
import { promisify } from "node:util";
import { RISK_LEVELS } from "../core/permissions.mjs";

const execAsync = promisify(exec);

export function createMediaDomain({ runtime }) {
  const capturesDir = path.join(process.cwd(), "data", "captures");

  async function ensureCapturesDir() {
    await fs.mkdir(capturesDir, { recursive: true });
  }

  return {
    name: "media",
    description: "Captura visual (escritorio, ventanas, aplicaciones) y gestión de impresión de documentos.",
    permissions: {
      capture_desktop: RISK_LEVELS.HIGH,
      capture_window: RISK_LEVELS.HIGH,
      capture_app: RISK_LEVELS.HIGH,
      list_printers: RISK_LEVELS.LOW,
      get_printer_info: RISK_LEVELS.LOW,
      print_document: RISK_LEVELS.HIGH,
      list_print_jobs: RISK_LEVELS.LOW,
      cancel_print_job: RISK_LEVELS.HIGH
    },
    actions: {
      async capture_desktop({ outputPath = null }) {
        await ensureCapturesDir();
        const targetFile = outputPath ? path.resolve(outputPath) : path.join(capturesDir, `screenshot_${Date.now()}.png`);

        const psScript = `
          Add-Type -AssemblyName System.Windows.Forms
          Add-Type -AssemblyName System.Drawing
          $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
          $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
          $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
          $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
          $bitmap.Save('${targetFile.replace(/\\/g, "\\\\")}', [System.Drawing.Imaging.ImageFormat]::Png)
          $graphics.Dispose()
          $bitmap.Dispose()
        `;

        await execAsync(`powershell.exe -NoProfile -Command "${psScript.replace(/"/g, '`"')}"`);
        const stat = await fs.stat(targetFile);

        return {
          ok: true,
          type: "desktop",
          path: targetFile,
          bytes: stat.size
        };
      },

      async capture_window({ query, outputPath = null }) {
        if (!query) throw new Error("Parámetro 'query' (título de ventana) es requerido.");
        await ensureCapturesDir();
        const targetFile = outputPath ? path.resolve(outputPath) : path.join(capturesDir, `window_${Date.now()}.png`);

        // Safe window capture via PowerShell parameter binding
        const psScript = `
          param([string]$targetQuery, [string]$outPath)
          Add-Type -AssemblyName System.Windows.Forms
          Add-Type -AssemblyName System.Drawing
          $proc = Get-Process | Where-Object { $_.MainWindowTitle -like "*$targetQuery*" -and $_.MainWindowHandle -ne 0 } | Select-Object -First 1
          if (-not $proc) { throw "Ventana con título '$targetQuery' no encontrada." }
          $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
          $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
          $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
          $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
          $bitmap.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
          $graphics.Dispose()
          $bitmap.Dispose()
        `;

        await execAsync(`powershell.exe -NoProfile -Command "${psScript.replace(/"/g, '`"')}" -targetQuery "${query.replace(/"/g, '`"')}" -outPath "${targetFile.replace(/"/g, '`"')}"`);
        const stat = await fs.stat(targetFile);
        return { ok: true, type: "window", query, path: targetFile, bytes: stat.size };
      },

      async capture_app({ processName, outputPath = null }) {
        if (!processName) throw new Error("Parámetro 'processName' requerido.");
        return this.capture_window({ query: processName, outputPath });
      },

      async list_printers() {
        try {
          const cmd = "Get-Printer -ErrorAction SilentlyContinue | Select-Object Name, Type, DriverName, PortName, PrinterStatus, Default | ConvertTo-Json";
          const { stdout } = await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
          let data = [];
          try {
            const parsed = JSON.parse(stdout);
            data = Array.isArray(parsed) ? parsed : [parsed];
          } catch {
            data = [];
          }
          return { ok: true, count: data.length, printers: data };
        } catch (e) {
          return { ok: true, count: 0, printers: [], warning: "Servicio de cola de impresión no disponible." };
        }
      },

      async get_printer_info({ name }) {
        if (!name) throw new Error("Parámetro 'name' de impresora requerido.");
        try {
          const cmd = `Get-Printer -Name "${name}" -ErrorAction SilentlyContinue | Select-Object * | ConvertTo-Json`;
          const { stdout } = await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
          return { ok: true, printer: JSON.parse(stdout) };
        } catch (e) {
          return { ok: false, error: `Impresora '${name}' no disponible o servicio detenido.` };
        }
      },

      async print_document({ filePath, printerName = null, copies = 1 }) {
        if (!filePath) throw new Error("Parámetro 'filePath' requerido.");
        const fullPath = path.resolve(filePath);
        await fs.access(fullPath);

        const safePath = fullPath.replace(/[`"$]/g, "");
        const safePrinter = printerName ? printerName.replace(/[`"$]/g, "") : null;
        const printerArg = safePrinter ? `-PrinterName "${safePrinter}"` : "";
        const cmd = `Start-Process -FilePath "${safePath}" -Verb PrintTo ${printerArg} -PassThru`;

        try {
          await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
          return { ok: true, file: fullPath, printer: safePrinter || "Default", copies };
        } catch (e) {
          // Fallback to Out-Printer
          const fallback = `Get-Content "${safePath}" | Out-Printer ${printerArg}`;
          await execAsync(`powershell.exe -NoProfile -Command "${fallback}"`);
          return { ok: true, file: fullPath, printer: safePrinter || "Default", fallback: true };
        }
      },

      async list_print_jobs({ printerName = null }) {
        let cmd = "Get-PrintJob";
        if (printerName) cmd += ` -PrinterName "${printerName}"`;
        cmd += " | Select-Object Id, DocumentName, JobStatus, TotalPages, SubmittedTime | ConvertTo-Json";

        const { stdout } = await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
        let data = [];
        try {
          const parsed = JSON.parse(stdout);
          data = Array.isArray(parsed) ? parsed : [parsed];
        } catch {
          data = [];
        }
        return { ok: true, count: data.length, jobs: data };
      },

      async cancel_print_job({ jobId, printerName }) {
        if (!jobId || !printerName) throw new Error("Parámetros 'jobId' y 'printerName' requeridos.");
        const cleanJobId = parseInt(jobId, 10);
        if (isNaN(cleanJobId)) throw new Error("Parámetro 'jobId' debe ser numérico.");
        const safePrinter = printerName.replace(/[`"$]/g, "");
        const cmd = `Remove-PrintJob -PrinterName "${safePrinter}" -ID ${cleanJobId}`;
        await execAsync(`powershell.exe -NoProfile -Command "${cmd}"`);
        return { ok: true, message: `Trabajo de impresión '${cleanJobId}' cancelado en '${safePrinter}'.` };
      }
    }
  };
}
