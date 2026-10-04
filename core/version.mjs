/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * core/version.mjs — Single Source of Truth Version Engine
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const versionFilePath = path.resolve(__dirname, "..", "VERSION");

let cachedVersion = null;

export function getVersion() {
  if (cachedVersion) return cachedVersion;
  try {
    cachedVersion = fs.readFileSync(versionFilePath, "utf8").trim();
  } catch {
    cachedVersion = "V26-0.1";
  }
  return cachedVersion;
}

export function getSemver(verStr = getVersion()) {
  const clean = String(verStr).trim().replace(/^v/i, "");
  // Format V26-0.1 -> 26.0.1
  const match = clean.match(/^(\d+)[-.](\d+)(?:\.(\d+))?$/);
  if (match) {
    const year = match[1];
    const major = match[2];
    const patch = match[3] ?? "0";
    return `${year}.${major}.${patch}`;
  }
  return clean;
}

/**
 * Compare two Fluxer XZ versions or semvers.
 * Ensures: V26-0.1 < V26-0.2 < V26-1.0 < V27-0.1
 * @param {string} a 
 * @param {string} b 
 * @returns {number} -1 if a < b, 0 if a == b, 1 if a > b
 */
export function compareVersions(a, b) {
  const parse = (v) => {
    const s = String(v).trim().replace(/^v/i, "");
    // Check if format is Year-Version (e.g. 26-0.1 or 26-1.0)
    const yearDash = s.match(/^(\d+)-(\d+)(?:\.(\d+))?$/);
    if (yearDash) {
      return [
        parseInt(yearDash[1], 10),
        parseInt(yearDash[2], 10),
        parseInt(yearDash[3] || "0", 10)
      ];
    }
    // Check if dot separated
    const parts = s.split(".").map((p) => parseInt(p, 10) || 0);
    while (parts.length < 3) parts.push(0);
    return parts;
  };

  const pa = parse(a);
  const pb = parse(b);

  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const na = pa[i] || 0;
    const nb = pb[i] || 0;
    if (na < nb) return -1;
    if (na > nb) return 1;
  }
  return 0;
}

export const PRODUCT_NAME = "FLUXER XZ - PUBLIC RELEASE";
export const FULL_VERSION_STRING = `${PRODUCT_NAME} ${getVersion()}`;
