/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * tools/web.mjs — High-Efficiency Web Search, Content Extraction & Downloads
 */
import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { compact } from "../core/compact.mjs";
import { RISK_LEVELS } from "../core/permissions.mjs";

export function createWebDomain({ runtime }) {
  return {
    name: "web",
    description: "Búsqueda web sin tokens externos, extracción de contenido y descargas seguras.",
    permissions: {
      search_web: RISK_LEVELS.LOW,
      extract_page: RISK_LEVELS.LOW,
      download_file: RISK_LEVELS.MEDIUM
    },
    actions: {
      async search_web({ query, limit = 5 }) {
        if (!query) throw new Error("Parámetro 'query' requerido.");
        const encoded = encodeURIComponent(query);
        const url = `https://html.duckduckgo.com/html/?q=${encoded}`;

        const res = await fetch(url, {
          headers: {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
          }
        });
        if (!res.ok) throw new Error(`HTTP ${res.status} al buscar en la web.`);
        const html = await res.text();

        const results = [];
        const resultRegex = /<a\s+class="result__url"\s+href="([^"]+)">[\s\S]*?<\/a>[\s\S]*?<a\s+class="result__snippet[^"]*"\s+href="[^"]*">([\s\S]*?)<\/a>/gi;
        let match;
        while ((match = resultRegex.exec(html)) !== null && results.length < limit) {
          const rawUrl = match[1].trim();
          const cleanSnippet = match[2].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
          results.push({
            url: rawUrl.startsWith("//") ? `https:${rawUrl}` : rawUrl,
            snippet: cleanSnippet
          });
        }

        return { ok: true, query, count: results.length, results };
      },

      async extract_page({ url, maxLength = 6000 }) {
        if (!url) throw new Error("Parámetro 'url' requerido.");
        const res = await fetch(url, {
          headers: {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
          }
        });
        if (!res.ok) throw new Error(`HTTP ${res.status} al descargar página.`);
        const html = await res.text();

        // Strip scripts, styles, head, tags
        let text = html
          .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
          .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "")
          .replace(/<head\b[^<]*(?:(?!<\/head>)<[^<]*)*<\/head>/gi, "")
          .replace(/<[^>]+>/g, " ")
          .replace(/&nbsp;/gi, " ")
          .replace(/&amp;/gi, "&")
          .replace(/&lt;/gi, "<")
          .replace(/&gt;/gi, ">")
          .replace(/\s+/g, " ")
          .trim();

        if (text.length > maxLength) {
          text = text.slice(0, maxLength) + "\n... [Texto truncado por Compact Mode] ...";
        }

        return { ok: true, url, length: text.length, content: text };
      },

      async download_file({ url, destPath, expectedSha256 = null }) {
        if (!url || !destPath) throw new Error("Parámetros 'url' y 'destPath' requeridos.");
        const resolvedDest = path.resolve(destPath);
        await fs.mkdir(path.dirname(resolvedDest), { recursive: true });

        const res = await fetch(url);
        if (!res.ok) throw new Error(`HTTP ${res.status} al descargar archivo.`);
        const buffer = Buffer.from(await res.arrayBuffer());

        const hash = crypto.createHash("sha256").update(buffer).digest("hex");
        if (expectedSha256 && expectedSha256.toLowerCase() !== hash.toLowerCase()) {
          throw new Error(`Checksum mismatch: esperado ${expectedSha256}, descargado ${hash}`);
        }

        await fs.writeFile(resolvedDest, buffer);
        return { ok: true, path: resolvedDest, bytes: buffer.length, sha256: hash };
      }
    }
  };
}