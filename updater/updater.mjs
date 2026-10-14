/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2
 * updater/updater.mjs — Autonomous Zero-Dependency Remote Updater Engine
 */
import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { exec } from "node:child_process";
import { fileURLToPath } from "node:url";
import { getVersion, compareVersions } from "../core/version.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");

export class UpdaterEngine {
  constructor(options = {}) {
    this.rootDir = options.rootDir || ROOT_DIR;
    this.repoOwner = options.repoOwner || "Lara2026ss";
    this.repoName = options.repoName || process.env.AERO_HOST_REPO_NAME || "aero-host-suite";
    this.apiBase = options.apiBase || "https://api.github.com";
    this.stagingDir = path.join(this.rootDir, "data", "staging");
    this.backupDir = path.join(this.rootDir, "data", "backups");
  }

  async ensureDirs() {
    await fs.mkdir(this.stagingDir, { recursive: true });
    await fs.mkdir(this.backupDir, { recursive: true });
  }

  computeFileHash(filePath) {
    return new Promise(async (resolve, reject) => {
      try {
        const data = await fs.readFile(filePath);
        const hash = crypto.createHash("sha256").update(data).digest("hex");
        resolve(hash);
      } catch (e) {
        reject(e);
      }
    });
  }

  async extractArchive(archivePath, targetDir) {
    await fs.mkdir(targetDir, { recursive: true });
    // Use Windows native tar.exe (supports zip and tar.gz)
    return new Promise((resolve, reject) => {
      exec(`tar.exe -xf "${archivePath}" -C "${targetDir}"`, (err, stdout, stderr) => {
        if (err) {
          // Fallback to powershell Expand-Archive
          exec(`powershell.exe -NoProfile -Command "Expand-Archive -LiteralPath '${archivePath}' -DestinationPath '${targetDir}' -Force"`, (err2) => {
            if (err2) return reject(new Error(`Failed to extract archive: ${stderr || err2.message}`));
            resolve(true);
          });
          return;
        }
        resolve(true);
      });
    });
  }

  async createBackupZip(destZipPath) {
    await this.ensureDirs();
    return new Promise((resolve, reject) => {
      // Archive root excluding data, node_modules, and .git using native tar.exe
      const cmd = `tar.exe -a -c -f "${destZipPath}" --exclude="data" --exclude="node_modules" --exclude=".git" *`;
      exec(cmd, { cwd: this.rootDir }, (err, stdout, stderr) => {
        if (err) return reject(new Error(`Backup creation failed: ${stderr || err.message}`));
        resolve(destZipPath);
      });
    });
  }

  async check(endpointOverride = null) {
    let currentVersion = getVersion();
    try {
      const vFile = path.join(this.rootDir, "VERSION");
      currentVersion = (await fs.readFile(vFile, "utf8")).trim();
    } catch {}
    let releaseData = null;

    if (endpointOverride) {
      if (typeof endpointOverride === "object") {
        releaseData = endpointOverride;
      } else {
        const res = await fetch(endpointOverride, {
          headers: { "User-Agent": "Aero-Host-Updater" }
        });
        if (!res.ok) throw new Error(`HTTP ${res.status} checking update`);
        releaseData = await res.json();
      }
    } else {
      const url = `${this.apiBase}/repos/${this.repoOwner}/${this.repoName}/releases/latest`;
      let res = await fetch(url, {
        headers: { "User-Agent": "Aero-Host-Updater" }
      });
      if (!res.ok) {
        // Fallback to /releases list (catches recently published releases before GitHub latest-tag indexer runs)
        const listUrl = `${this.apiBase}/repos/${this.repoOwner}/${this.repoName}/releases`;
        const listRes = await fetch(listUrl, {
          headers: { "User-Agent": "Aero-Host-Updater" }
        });
        if (listRes.ok) {
          const list = await listRes.json();
          if (Array.isArray(list) && list.length > 0) {
            releaseData = list[0];
          }
        }
        if (!releaseData) {
          return { ok: true, hasUpdate: false, currentVersion, message: "No releases found on repository." };
        }
      } else {
        releaseData = await res.json();
      }
    }

    const latestTag = releaseData.tag_name || releaseData.tagName || releaseData.name || "";
    const cleanTag = latestTag.replace(/^v/i, "V");
    const hasUpdate = compareVersions(currentVersion, cleanTag) < 0;

    // Find asset package and checksum
    const assets = releaseData.assets || [];
    const pkgAsset = assets.find((a) => a.name.endsWith(".zip") || a.name.endsWith(".tar.gz"));
    const shaAsset = assets.find((a) => a.name.endsWith(".sha256") || a.name.endsWith(".checksum"));

    return {
      ok: true,
      hasUpdate,
      currentVersion,
      latestVersion: cleanTag,
      publishedAt: releaseData.published_at || releaseData.publishedAt,
      releaseUrl: releaseData.html_url || releaseData.htmlUrl,
      notes: releaseData.body || "",
      packageUrl: pkgAsset?.browser_download_url || releaseData.packageUrl || null,
      expectedSha256: releaseData.expectedSha256 || null,
      shaUrl: shaAsset?.browser_download_url || null
    };
  }

  async apply(options = {}) {
    await this.ensureDirs();
    const currentVersion = getVersion();
    const checkRes = options.releaseOverride?.latestVersion ? options.releaseOverride : await this.check(options.releaseOverride);

    if (!checkRes.hasUpdate && !options.force) {
      return { ok: true, applied: false, message: `Ya está en la versión más reciente (${currentVersion}).` };
    }

    const targetVersion = checkRes.latestVersion;
    const pkgUrl = options.packageUrl || checkRes.packageUrl;
    if (!pkgUrl) {
      throw new Error(`No se encontró el archivo de paquete para la versión ${targetVersion}`);
    }

    // 1. Download to staging
    const pkgFilename = `aero-host-${targetVersion}.zip`;
    const stagingPkgPath = path.join(this.stagingDir, pkgFilename);

    if (options.localPackagePath) {
      await fs.copyFile(options.localPackagePath, stagingPkgPath);
    } else {
      const resp = await fetch(pkgUrl);
      if (!resp.ok) throw new Error(`HTTP ${resp.status} al descargar paquete`);
      const buffer = Buffer.from(await resp.arrayBuffer());
      await fs.writeFile(stagingPkgPath, buffer);
    }

    // 2. Verify SHA-256
    const localHash = await this.computeFileHash(stagingPkgPath);
    let expectedHash = options.expectedSha256 || checkRes.expectedSha256;
    if (!expectedHash && checkRes.shaUrl) {
      const shaResp = await fetch(checkRes.shaUrl);
      if (shaResp.ok) {
        const text = await shaResp.text();
        expectedHash = text.trim().split(/\s+/)[0];
      }
    }

    if (expectedHash) {
      const normExpected = expectedHash.trim().toLowerCase();
      const normLocal = localHash.trim().toLowerCase();
      if (normExpected !== normLocal) {
        await fs.unlink(stagingPkgPath).catch(() => {});
        throw new Error(`CHECKSUM_MISMATCH: SHA-256 esperado (${normExpected}) no coincide con el descargado (${normLocal})`);
      }
    }

    // 3. Create Backup of Current Version
    const backupName = `backup_${currentVersion}_${Date.now()}.zip`;
    const backupPath = path.join(this.backupDir, backupName);
    await this.createBackupZip(backupPath);

    // 4. Extract into staging directory
    const extractDir = path.join(this.stagingDir, `extracted_${targetVersion}_${Date.now()}`);
    await this.extractArchive(stagingPkgPath, extractDir);

    // 5. Replace files atomically (skipping data/ and user config)
    try {
      await this.atomicCopy(extractDir, this.rootDir);
    } catch (copyErr) {
      // Automatic Rollback!
      await this.rollback({ backupPath });
      throw new Error(`Error durante reemplazo de archivos. Rollback ejecutado: ${copyErr.message}`);
    }

    // 6. VERIFY PHYSICAL DISK CHANGE
    const diskVersionPath = path.join(this.rootDir, "VERSION");
    let newDiskVersion = "";
    try {
      newDiskVersion = (await fs.readFile(diskVersionPath, "utf8")).trim();
    } catch {}

    if (compareVersions(newDiskVersion, targetVersion) !== 0 && !options.skipVersionCheck) {
      // Disk verification failed! Rollback!
      await this.rollback({ backupPath });
      throw new Error(`VERIFICATION_FAILED: El archivo VERSION en disco (${newDiskVersion}) no coincide con la versión esperada (${targetVersion}). Rollback automático ejecutado.`);
    }

    // Clean staging
    await fs.rm(extractDir, { recursive: true, force: true }).catch(() => {});
    await fs.unlink(stagingPkgPath).catch(() => {});

    return {
      ok: true,
      applied: true,
      previousVersion: currentVersion,
      newVersion: targetVersion,
      backupPath,
      restartRequired: true,
      message: `Actualización exitosa a ${targetVersion}. Backup guardado en ${path.basename(backupPath)}. Reinicio de cliente MCP requerido.`
    };
  }

  async atomicCopy(srcDir, destDir) {
    const entries = await fs.readdir(srcDir, { withFileTypes: true });
    for (const entry of entries) {
      // STRICT PROTECTION: Never overwrite data/, node_modules/ or .git
      if (["data", "node_modules", ".git"].includes(entry.name)) continue;

      const srcPath = path.join(srcDir, entry.name);
      const destPath = path.join(destDir, entry.name);

      if (entry.isDirectory()) {
        await fs.mkdir(destPath, { recursive: true });
        await this.atomicCopy(srcPath, destPath);
      } else {
        await fs.copyFile(srcPath, destPath);
      }
    }
  }

  async rollback(options = {}) {
    await this.ensureDirs();
    let targetBackup = options.backupPath;

    if (!targetBackup) {
      const files = (await fs.readdir(this.backupDir))
        .filter((f) => f.startsWith("backup_") && f.endsWith(".zip"))
        .sort()
        .reverse();

      if (files.length === 0) {
        throw new Error("No se encontraron copias de seguridad en data/backups para rollback.");
      }
      targetBackup = path.join(this.backupDir, files[0]);
    }

    // Extract backup to root
    await this.extractArchive(targetBackup, this.rootDir);

    // Verify disk version after rollback
    const diskVersion = (await fs.readFile(path.join(this.rootDir, "VERSION"), "utf8")).trim();

    return {
      ok: true,
      restoredVersion: diskVersion,
      backupUsed: targetBackup,
      message: `Rollback completado exitosamente a versión ${diskVersion}.`
    };
  }

  async getStatus() {
    await this.ensureDirs();
    const backups = (await fs.readdir(this.backupDir)).filter((f) => f.endsWith(".zip"));
    return {
      ok: true,
      version: getVersion(),
      status: "healthy",
      rootDir: this.rootDir,
      backupsCount: backups.length,
      latestBackup: backups.sort().reverse()[0] || null
    };
  }
}

// Standalone CLI execution
if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const engine = new UpdaterEngine();
  const arg = process.argv[2] || "--status";

  (async () => {
    try {
      if (arg === "--check") {
        console.log(JSON.stringify(await engine.check(), null, 2));
      } else if (arg === "--apply") {
        console.log(JSON.stringify(await engine.apply(), null, 2));
      } else if (arg === "--rollback") {
        console.log(JSON.stringify(await engine.rollback(), null, 2));
      } else {
        console.log(JSON.stringify(await engine.getStatus(), null, 2));
      }
    } catch (e) {
      console.error(e.message);
      process.exit(1);
    }
  })();
}
