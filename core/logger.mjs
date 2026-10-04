/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * core/logger.mjs — High-Performance Structured Logger
 */
import fs from "node:fs/promises";
import path from "node:path";

export class Logger {
  constructor(options = {}) {
    this.rootDir = options.rootDir || process.cwd();
    this.logFile = path.join(this.rootDir, "data", "logs", "fluxer.log");
  }

  async log(level, message, meta = {}) {
    const entry = {
      timestamp: new Date().toISOString(),
      level: String(level).toUpperCase(),
      message,
      ...(Object.keys(meta).length ? { meta } : {})
    };

    try {
      await fs.mkdir(path.dirname(this.logFile), { recursive: true });
      await fs.appendFile(this.logFile, JSON.stringify(entry) + "\n", "utf8");
    } catch {}

    if (level === "error" || process.env.FLUXER_DEBUG) {
      process.stderr.write(`[FLUXER ${entry.level}] ${message}\n`);
    }
  }

  info(msg, meta) { return this.log("info", msg, meta); }
  warn(msg, meta) { return this.log("warn", msg, meta); }
  error(msg, meta) { return this.log("error", msg, meta); }
  debug(msg, meta) { return this.log("debug", msg, meta); }
}

export const logger = new Logger();
