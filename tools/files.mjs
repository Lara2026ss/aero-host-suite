/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
 * tools/files.mjs — High-Efficiency File System Operations with Compact Mode
 */
import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { compact } from "../core/compact.mjs";
import { RISK_LEVELS } from "../core/permissions.mjs";

export function createFilesDomain({ runtime }) {
  const trashDir = path.join(process.cwd(), "data", "trash");

  async function ensureTrash() {
    await fs.mkdir(trashDir, { recursive: true });
  }

  return {
    name: "files",
    description: "Operaciones de archivos: lectura por rango, escritura incremental, parches, búsqueda y gestión.",
    permissions: {
      read_file: RISK_LEVELS.LOW,
      write_file: RISK_LEVELS.MEDIUM,
      patch_file: RISK_LEVELS.MEDIUM,
      list_directory: RISK_LEVELS.LOW,
      search_files: RISK_LEVELS.LOW,
      copy_path: RISK_LEVELS.MEDIUM,
      move_path: RISK_LEVELS.MEDIUM,
      delete_path: RISK_LEVELS.DESTRUCTIVE,
      get_file_info: RISK_LEVELS.LOW
    },
    actions: {
      async read_file({ path: targetPath, handle, startLine = 1, endLine = null, checkUnchanged = true, detail = false }) {
        let resolvedPath = targetPath;
        if (handle) {
          const h = compact.resolveHandle(handle);
          if (!h) throw new Error(`Handle '${handle}' no encontrado o caducado.`);
          resolvedPath = h.path;
        }
        if (!resolvedPath) throw new Error("Parámetro 'path' o 'handle' requerido.");

        const fullPath = path.resolve(resolvedPath);
        const content = await fs.readFile(fullPath, "utf8");

        // Hash unchanged check (only when reading entire file)
        if (checkUnchanged && startLine === 1 && endLine === null && compact.getMode() !== "off") {
          const check = compact.checkUnchanged(fullPath, content);
          if (check.unchanged) {
            return {
              ok: true,
              unchanged: true,
              hash: check.hash,
              message: "Sin cambios desde la última lectura."
            };
          }
        }

        const hId = compact.createHandle("file", fullPath);

        // Range read
        if (startLine > 1 || endLine !== null) {
          const slice = compact.readByRange(content, startLine, endLine);
          return {
            ok: true,
            handle: hId,
            startLine: slice.start,
            endLine: slice.end,
            totalLines: slice.totalLines,
            lineCount: slice.lineCount,
            content: slice.content
          };
        }

        const lines = content.split(/\r?\n/).length;
        return {
          ok: true,
          handle: hId,
          totalLines: lines,
          content: detail || lines < 100 ? content : compact.truncateOutput(content, 80)
        };
      },

      async write_file({ path: targetPath, content, stashId, createDirs = true }) {
        if (!targetPath) throw new Error("Parámetro 'path' es requerido.");
        let payload = content;
        if (stashId) {
          payload = compact.stashGet(stashId);
          if (payload === null) throw new Error(`StashId '${stashId}' no encontrado.`);
        }
        if (typeof payload !== "string") throw new Error("Contenido no especificado (use 'content' o 'stashId').");

        const fullPath = path.resolve(targetPath);
        if (createDirs) {
          await fs.mkdir(path.dirname(fullPath), { recursive: true });
        }

        await fs.writeFile(fullPath, payload, "utf8");
        const hash = compact.computeHash(payload);
        compact.recordRead(fullPath, payload);

        return {
          ok: true,
          path: fullPath,
          bytesWritten: Buffer.byteLength(payload, "utf8"),
          hash
        };
      },

      async patch_file({ path: targetPath, target, replacement, startLine, endLine, allowMultiple = false }) {
        if (!targetPath) throw new Error("Parámetro 'path' requerido.");
        const fullPath = path.resolve(targetPath);
        const original = await fs.readFile(fullPath, "utf8");

        let patchResult;
        if (target !== undefined) {
          patchResult = compact.applyAnchorPatch(original, { target, replacement, allowMultiple });
        } else if (startLine !== undefined && endLine !== undefined) {
          patchResult = compact.applyLineRangePatch(original, { startLine, endLine, replacement });
        } else {
          throw new Error("Especifique 'target' (para parche por ancla) o 'startLine'/'endLine' (para parche por rango).");
        }

        await fs.writeFile(fullPath, patchResult.modified, "utf8");
        compact.recordRead(fullPath, patchResult.modified);

        return {
          ok: true,
          path: fullPath,
          beforeHash: patchResult.beforeHash,
          afterHash: patchResult.afterHash,
          replacements: patchResult.replacements || patchResult.replacedLines
        };
      },

      async list_directory({ path: dirPath = ".", recursive = false, limit = 100 }) {
        const fullDir = path.resolve(dirPath);
        const entries = await fs.readdir(fullDir, { withFileTypes: true });
        const results = [];

        for (const entry of entries) {
          if (["node_modules", ".git"].includes(entry.name)) continue;
          const entryFull = path.join(fullDir, entry.name);
          let size = 0;
          try {
            if (entry.isFile()) {
              const stat = await fs.stat(entryFull);
              size = stat.size;
            }
          } catch {}

          results.push({
            name: entry.name,
            type: entry.isDirectory() ? "directory" : "file",
            size
          });

          if (results.length >= limit) break;
        }

        return {
          ok: true,
          directory: fullDir,
          count: results.length,
          entries: results
        };
      },

      async search_files({ dir = ".", query, limit = 50 }) {
        if (!query) throw new Error("Parámetro 'query' requerido.");
        const fullDir = path.resolve(dir);
        const results = [];

        async function walk(current) {
          if (results.length >= limit) return;
          const entries = await fs.readdir(current, { withFileTypes: true }).catch(() => []);
          for (const e of entries) {
            if (["node_modules", ".git", "data"].includes(e.name)) continue;
            const itemPath = path.join(current, e.name);
            if (e.name.toLowerCase().includes(query.toLowerCase())) {
              results.push({ name: e.name, path: itemPath, isDirectory: e.isDirectory() });
              if (results.length >= limit) return;
            }
            if (e.isDirectory()) {
              await walk(itemPath);
            }
          }
        }

        await walk(fullDir);
        return { ok: true, query, count: results.length, matches: results };
      },

      async copy_path({ src, dest }) {
        if (!src || !dest) throw new Error("Parámetros 'src' y 'dest' requeridos.");
        const resolvedSrc = path.resolve(src);
        const resolvedDest = path.resolve(dest);
        await fs.mkdir(path.dirname(resolvedDest), { recursive: true });
        await fs.cp(resolvedSrc, resolvedDest, { recursive: true });
        return { ok: true, src: resolvedSrc, dest: resolvedDest };
      },

      async move_path({ src, dest }) {
        if (!src || !dest) throw new Error("Parámetros 'src' y 'dest' requeridos.");
        const resolvedSrc = path.resolve(src);
        const resolvedDest = path.resolve(dest);
        await fs.mkdir(path.dirname(resolvedDest), { recursive: true });
        try {
          await fs.rename(resolvedSrc, resolvedDest);
        } catch (renameErr) {
          if (renameErr.code === "EXDEV") {
            await fs.cp(resolvedSrc, resolvedDest, { recursive: true });
            await fs.rm(resolvedSrc, { recursive: true, force: true });
          } else {
            throw renameErr;
          }
        }
        return { ok: true, src: resolvedSrc, dest: resolvedDest };
      },

      async delete_path({ path: targetPath, permanent = false }) {
        if (!targetPath) throw new Error("Parámetro 'path' requerido.");
        const fullPath = path.resolve(targetPath);
        const stat = await fs.stat(fullPath);

        if (!permanent) {
          await ensureTrash();
          const timestamp = Date.now();
          const trashTarget = path.join(trashDir, `${path.basename(fullPath)}_${timestamp}`);
          try {
            await fs.rename(fullPath, trashTarget);
          } catch (renameErr) {
            if (renameErr.code === "EXDEV") {
              await fs.cp(fullPath, trashTarget, { recursive: true });
              await fs.rm(fullPath, { recursive: true, force: true });
            } else {
              throw renameErr;
            }
          }
          return { ok: true, movedToTrash: true, trashPath: trashTarget };
        }

        if (stat.isDirectory()) {
          await fs.rm(fullPath, { recursive: true, force: true });
        } else {
          await fs.unlink(fullPath);
        }
        return { ok: true, deleted: true, path: fullPath };
      },

      async get_file_info({ path: targetPath }) {
        if (!targetPath) throw new Error("Parámetro 'path' requerido.");
        const fullPath = path.resolve(targetPath);
        const stat = await fs.stat(fullPath);
        let hash = null;
        let lineCount = null;

        if (stat.isFile()) {
          const content = await fs.readFile(fullPath);
          hash = crypto.createHash("sha256").update(content).digest("hex");
          if (stat.size < 10 * 1024 * 1024) {
            lineCount = content.toString("utf8").split(/\r?\n/).length;
          }
        }

        return {
          ok: true,
          path: fullPath,
          size: stat.size,
          isDirectory: stat.isDirectory(),
          isFile: stat.isFile(),
          createdAt: stat.birthtime,
          modifiedAt: stat.mtime,
          hash,
          lineCount
        };
      }
    }
  };
}
