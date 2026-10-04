/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * tools/feedback.mjs — Privacy-Safe Feedback and GitHub Issue Bridge
 */
import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import os from "node:os";
import { getVersion } from "../core/version.mjs";
import { RISK_LEVELS } from "../core/permissions.mjs";

export function createFeedbackDomain({ runtime }) {
  const feedbackDir = path.resolve(runtime?.dirs?.feedback || path.join(process.cwd(), "data", "feedback"));
  const officialRepo = "Lara2026ss/aero-fluxer-x";

  async function ensureDir() {
    await fs.mkdir(feedbackDir, { recursive: true });
  }

  function sanitizeContext(ctx) {
    if (!ctx || typeof ctx !== "object") return {};
    const sanitized = {};
    const userHome = os.homedir();
    for (const [k, v] of Object.entries(ctx)) {
      if (typeof v === "string") {
        sanitized[k] = v.replaceAll(userHome, "~").replace(/C:\\Users\\[^\\]+/gi, "~");
      } else {
        sanitized[k] = v;
      }
    }
    return sanitized;
  }

  function generateIssueUrl(feedback) {
    const title = encodeURIComponent(`[${feedback.type.toUpperCase()}] ${feedback.tool ? `[${feedback.tool}] ` : ""}${feedback.description.slice(0, 60)}`);
    const body = encodeURIComponent(
`### Tipo de Feedback
${feedback.type}

### Herramienta / Acción
${feedback.tool || "General"}${feedback.action ? `.${feedback.action}` : ""}

### Severidad
${feedback.severity}

### Descripción
${feedback.description}

### Contexto del Sistema
- **Versión Fluxer XZ**: ${feedback.version}
- **Sistema Operativo**: ${feedback.platform} (${feedback.arch})
- **Fecha**: ${feedback.createdAt}
${Object.keys(feedback.context || {}).length ? `\n### Datos Adicionales (Anonimizados)\n\`\`\`json\n${JSON.stringify(feedback.context, null, 2)}\n\`\`\`` : ""}
`
    );
    return `https://github.com/${officialRepo}/issues/new?title=${title}&body=${body}`;
  }

  return {
    name: "feedback",
    description: "Gestión de retroalimentación, reporte de errores y sugerencias de la IA o el usuario.",
    permissions: {
      send_feedback: RISK_LEVELS.LOW,
      list_feedback: RISK_LEVELS.LOW,
      view_feedback: RISK_LEVELS.LOW,
      update_status: RISK_LEVELS.LOW,
      delete_feedback: RISK_LEVELS.LOW,
      export_feedback: RISK_LEVELS.LOW,
      open_github_issue: RISK_LEVELS.LOW
    },
    actions: {
      async send_feedback({ type = "sugerencia", tool = "", action = "", description, severity = "medium", context = {} }) {
        if (!description) throw new Error("La descripción del feedback es obligatoria.");
        await ensureDir();

        const id = `fb_${Date.now()}_${crypto.randomBytes(3).toString("hex")}`;
        const feedback = {
          id,
          type,
          tool,
          action,
          description,
          severity,
          status: "abierto",
          version: getVersion(),
          platform: `${os.type()} ${os.release()}`,
          arch: os.arch(),
          context: sanitizeContext(context),
          createdAt: new Date().toISOString()
        };

        const filePath = path.join(feedbackDir, `${id}.json`);
        await fs.writeFile(filePath, JSON.stringify(feedback, null, 2), "utf8");

        const githubIssueUrl = generateIssueUrl(feedback);

        return {
          ok: true,
          id,
          message: "Feedback guardado localmente con éxito.",
          githubIssueUrl,
          instructions: "Para compartirlo públicamente con el mantenedor, abra la URL de GitHub issues generada y confirme con su cuenta de GitHub."
        };
      },

      async list_feedback({ type, status, limit = 50 }) {
        await ensureDir();
        const files = (await fs.readdir(feedbackDir)).filter((f) => f.endsWith(".json"));
        const list = [];

        for (const file of files) {
          try {
            const raw = await fs.readFile(path.join(feedbackDir, file), "utf8");
            const item = JSON.parse(raw);
            if (type && item.type !== type) continue;
            if (status && item.status !== status) continue;
            list.push({
              id: item.id,
              type: item.type,
              tool: item.tool,
              severity: item.severity,
              status: item.status,
              createdAt: item.createdAt,
              summary: item.description.slice(0, 80)
            });
            if (list.length >= limit) break;
          } catch {}
        }

        return { ok: true, count: list.length, feedbacks: list };
      },

      async view_feedback({ id }) {
        if (!id) throw new Error("ID de feedback requerido.");
        const filePath = path.join(feedbackDir, `${id}.json`);
        const raw = await fs.readFile(filePath, "utf8");
        const item = JSON.parse(raw);
        return { ok: true, feedback: item, githubIssueUrl: generateIssueUrl(item) };
      },

      async update_status({ id, status }) {
        if (!id || !status) throw new Error("ID y status son requeridos.");
        const filePath = path.join(feedbackDir, `${id}.json`);
        const raw = await fs.readFile(filePath, "utf8");
        const item = JSON.parse(raw);
        item.status = status;
        item.updatedAt = new Date().toISOString();
        await fs.writeFile(filePath, JSON.stringify(item, null, 2), "utf8");
        return { ok: true, message: `Feedback '${id}' actualizado a estado '${status}'.`, feedback: item };
      },

      async delete_feedback({ id }) {
        if (!id) throw new Error("ID de feedback requerido.");
        const filePath = path.join(feedbackDir, `${id}.json`);
        await fs.unlink(filePath);
        return { ok: true, message: `Feedback '${id}' eliminado correctamente.` };
      },

      async export_feedback({ id = null }) {
        await ensureDir();
        if (id) {
          const filePath = path.join(feedbackDir, `${id}.json`);
          const raw = await fs.readFile(filePath, "utf8");
          return { ok: true, data: JSON.parse(raw) };
        }
        const files = (await fs.readdir(feedbackDir)).filter((f) => f.endsWith(".json"));
        const all = [];
        for (const f of files) {
          try {
            const raw = await fs.readFile(path.join(feedbackDir, f), "utf8");
            all.push(JSON.parse(raw));
          } catch {}
        }
        return { ok: true, count: all.length, exportData: all };
      },

      async open_github_issue({ id }) {
        if (!id) throw new Error("ID de feedback requerido.");
        const filePath = path.join(feedbackDir, `${id}.json`);
        const raw = await fs.readFile(filePath, "utf8");
        const item = JSON.parse(raw);
        const url = generateIssueUrl(item);
        return { ok: true, githubIssueUrl: url, instructions: "Abra la URL en su navegador web para crear el issue en GitHub." };
      }
    }
  };
}
