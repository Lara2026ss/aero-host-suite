/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
 * core/config.mjs — Central Configuration Loader
 */
import fs from "node:fs/promises";
import path from "node:path";
import { getVersion } from "./version.mjs";

const DEFAULTS = {
  version: getVersion(),
  security: {
    title: "Aero Host Suite · Host Control",
    allowElevation: true,
    defaultRiskAction: "allow"
  },
  compact: {
    mode: "auto",
    maxChars: 8000,
    maxLines: 80
  },
  shortcuts: {
    folder: null // Dynamic, defaults to data/shortcuts
  },
  logging: {
    level: "info",
    auditEnabled: true
  }
};

export async function loadConfig(rootDir = process.cwd()) {
  const configPath = path.join(rootDir, "data", "config.json");
  let userConfig = {};

  try {
    const raw = await fs.readFile(configPath, "utf8");
    userConfig = JSON.parse(raw);
  } catch {}

  return {
    ...DEFAULTS,
    ...userConfig,
    security: { ...DEFAULTS.security, ...(userConfig.security || {}) },
    compact: { ...DEFAULTS.compact, ...(userConfig.compact || {}) }
  };
}

export async function saveConfig(config, rootDir = process.cwd()) {
  const configPath = path.join(rootDir, "data", "config.json");
  await fs.mkdir(path.dirname(configPath), { recursive: true });
  await fs.writeFile(configPath, JSON.stringify(config, null, 2), "utf8");
  return config;
}
