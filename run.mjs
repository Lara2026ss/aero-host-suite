#!/usr/bin/env node
/**
 * Aero Host Suite — Official Release V26-0.2
 * run.mjs — Direct CLI & Headless Runner for Aero Host Suite Capabilities
 * Enables any agent, script or automation to execute Aero Host Suite domains directly.
 */
import { fileURLToPath } from "node:url";
import path from "node:path";
import fs from "node:fs/promises";
import { Runtime } from "./core/runtime.mjs";

const ROOT_DIR = path.dirname(fileURLToPath(import.meta.url));

async function main() {
  const args = process.argv.slice(2);
  const toolName = args[0];
  let rawParams = args[1] || "{}";

  if (!toolName) {
    console.error(JSON.stringify({ ok: false, error: "Usage: node run.mjs <domain> <action|json_params> [json_args]" }));
    process.exit(1);
  }

  let action = "";
  let mergedParams = {};

  function parseTolerantJson(str) {
    if (!str || typeof str !== "string") return {};
    const trimmed = str.trim();
    if (!trimmed) return {};
    try {
      return JSON.parse(trimmed);
    } catch {}
    try {
      // Fix PowerShell stripping quotes from keys or values
      const fixed = trimmed
        .replace(/([{,]\s*)([a-zA-Z0-9_]+)\s*:/g, '$1"$2":')
        .replace(/'/g, '"');
      return JSON.parse(fixed);
    } catch {}
    return null;
  }

  // Check if second arg is an action name instead of JSON
  if (rawParams && !rawParams.startsWith("{") && !rawParams.endsWith(".json")) {
    action = rawParams;
    const thirdArg = args[2] || "{}";
    mergedParams = parseTolerantJson(thirdArg) || {};
  } else {
    // If second arg is a file path that exists, read it
    if (rawParams.endsWith(".json")) {
      try {
        rawParams = await fs.readFile(rawParams, "utf8");
      } catch {}
    }

    let params = parseTolerantJson(rawParams);
    if (!params) {
      console.error(JSON.stringify({ ok: false, error: "Invalid JSON params format" }));
      process.exit(1);
    }

    action = params.action;
    if (!action) {
      console.error(JSON.stringify({ ok: false, error: "Parameter 'action' is required" }));
      process.exit(1);
    }

    mergedParams = {
      ...(typeof params.args === "object" && params.args !== null ? params.args : {}),
      ...params,
    };
    delete mergedParams.args;
  }

  const runtime = new Runtime({ rootDir: ROOT_DIR });
  await runtime.init();
  runtime.security.setElevated(true, "*");


  try {
    const result = await runtime.router.execute(toolName, action, mergedParams);
    console.log(typeof result === "string" ? result : JSON.stringify(result, null, 2));
    if (result && result.ok === false) process.exit(1);
  } catch (err) {
    console.error(JSON.stringify({ ok: false, error: err.message, tool: toolName, action }));
    process.exit(1);
  }
}

main();
