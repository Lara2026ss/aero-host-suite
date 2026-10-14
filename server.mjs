#!/usr/bin/env node
/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * server.mjs — High-Performance Model Context Protocol (MCP) Server
 */
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { Runtime } from "./core/runtime.mjs";
import { getVersion, PRODUCT_NAME } from "./core/version.mjs";

const ROOT_DIR = path.dirname(fileURLToPath(import.meta.url));

export async function startServer() {
  const runtime = new Runtime({ rootDir: ROOT_DIR });
  await runtime.init();

  const mcpServer = new Server(
    {
      name: "Aero_Host_Suite",
      version: getVersion(),
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // 1. List Available Tools
  mcpServer.setRequestHandler(ListToolsRequestSchema, async () => {
    const tools = [];
    for (const toolName of runtime.registry.moduleNames()) {
      const mod = runtime.registry.get(toolName);
      const actions = Object.keys(mod.actions || {}).sort();

      tools.push({
        name: toolName,
        description: `Dominio '${toolName}': ${mod.description}\nAcciones disponibles: ${actions.join(", ")}`,
        inputSchema: {
          type: "object",
          properties: {
            action: {
              type: "string",
              enum: actions,
              description: `Acción a ejecutar en '${toolName}'.`,
            },
            path: { type: "string", description: "Ruta de archivo o carpeta." },
            command: { type: "string", description: "Comando o script de terminal." },
            query: { type: "string", description: "Término de búsqueda o consulta." },
            name: { type: "string", description: "Nombre de recurso, paquete o atajo." },
            content: { type: "string", description: "Contenido de texto o datos." },
            stashId: { type: "string", description: "Identificador de contenido previamente guardado." },
            handle: { type: "string", description: "Identificador corto de archivo o documento (f#1, doc#1)." },
            startLine: { type: "integer", description: "Línea inicial para lectura o reemplazo." },
            endLine: { type: "integer", description: "Línea final para lectura o reemplazo." },
            target: { type: "string", description: "Texto ancla exacto a reemplazar." },
            replacement: { type: "string", description: "Nuevo texto de reemplazo." },
            args: { type: "object", description: "Argumentos adicionales de la acción." },
          },
          required: ["action"],
          additionalProperties: true,
        },
      });
    }

    return { tools };
  });

  // 2. Call Tool Execution Handler
  mcpServer.setRequestHandler(CallToolRequestSchema, async (request) => {
    const toolName = request.params.name;
    const rawArgs = request.params.arguments || {};
    const action = rawArgs.action;

    if (!action) {
      return {
        isError: true,
        content: [
          {
            type: "text",
            text: JSON.stringify({
              ok: false,
              error: `Se requiere el parámetro 'action' para invocar la herramienta '${toolName}'.`,
            }),
          },
        ],
      };
    }

    // Merge nested 'args' with root properties
    const params = {
      ...(typeof rawArgs.args === "object" && rawArgs.args !== null ? rawArgs.args : {}),
      ...rawArgs,
    };
    delete params.args;

    try {
      const result = await runtime.router.execute(toolName, action, params);
      const isError = result.ok === false;
      const formattedText = typeof result === "string" ? result : JSON.stringify(result, null, 2);

      return {
        isError,
        content: [
          {
            type: "text",
            text: formattedText,
          },
        ],
      };
    } catch (err) {
      return {
        isError: true,
        content: [
          {
            type: "text",
            text: JSON.stringify({
              ok: false,
              error: err.message,
              tool: toolName,
              action,
            }),
          },
        ],
      };
    }
  });

  // Connect stdio transport
  const transport = new StdioServerTransport();
  await mcpServer.connect(transport);

  // Graceful shutdown
  process.on("SIGINT", () => { process.exit(0); });
  process.on("SIGTERM", () => { process.exit(0); });

  return mcpServer;
}

// Auto-run if executed directly
if (process.argv[1] === fileURLToPath(import.meta.url)) {
  startServer().catch((err) => {
    process.stderr.write(`[AERO HOST SUITE FATAL] ${err.stack || err.message}\n`);
    process.exit(1);
  });
}
