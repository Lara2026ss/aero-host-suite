/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * core/runtime.mjs — Central Runtime and Lifecycle Engine
 */
import path from "node:path";
import fs from "node:fs/promises";
import { Registry } from "./registry.mjs";
import { Router } from "./router.mjs";
import { security } from "./permissions.mjs";
import { compact } from "./compact.mjs";
import { getVersion, PRODUCT_NAME } from "./version.mjs";

export class Runtime {
  constructor(options = {}) {
    this.rootDir = path.resolve(options.rootDir || process.cwd());
    this.version = getVersion();
    this.product = PRODUCT_NAME;

    this.dirs = {
      root: this.rootDir,
      data: path.join(this.rootDir, "data"),
      shortcuts: path.join(this.rootDir, "data", "shortcuts"),
      logs: path.join(this.rootDir, "data", "logs"),
      staging: path.join(this.rootDir, "data", "staging"),
      backups: path.join(this.rootDir, "data", "backups"),
      feedback: path.join(this.rootDir, "data", "feedback"),
      security: path.join(this.rootDir, "data", "security")
    };

    this.security = security;
    this.compact = compact;
    this.registry = new Registry(this);
    this.router = new Router(this);
  }

  async init() {
    for (const d of Object.values(this.dirs)) {
      await fs.mkdir(d, { recursive: true });
    }
    const toolsLoaded = await this.registry.load();
    return {
      ok: true,
      product: this.product,
      version: this.version,
      toolsLoaded
    };
  }
}

export const runtime = new Runtime();
