/**
 * Aero Host Suite — tools/documents/templates.mjs
 * Professional Design Specifications & Template Presets for Office Automation
 * Supports Word, Excel, and PowerPoint structured generative workflows.
 */

export const THEMES = {
  modern_blue: {
    primary: "#1F4E79",
    secondary: "#2F5597",
    accent: "#5B9BD5",
    background: "#F2F4F7",
    cardBg: "#FFFFFF",
    textDark: "#1A1A1A",
    textLight: "#FFFFFF",
    border: "#D9D9D9",
    fontFamily: "Calibri",
  },
  politecnico_maroon: {
    primary: "#6C1D45",
    secondary: "#8A2459",
    accent: "#B85D8A",
    background: "#FDF8F9",
    cardBg: "#FFFFFF",
    textDark: "#2B2B2B",
    textLight: "#FFFFFF",
    border: "#EAD5DF",
    fontFamily: "Georgia",
  },
  emerald_executive: {
    primary: "#0E5A44",
    secondary: "#198754",
    accent: "#38B48B",
    background: "#F4F9F6",
    cardBg: "#FFFFFF",
    textDark: "#1D2A24",
    textLight: "#FFFFFF",
    border: "#D1E7DD",
    fontFamily: "Segoe UI",
  },
  slate_minimal: {
    primary: "#2B303A",
    secondary: "#475263",
    accent: "#76869E",
    background: "#F7F8FA",
    cardBg: "#FFFFFF",
    textDark: "#1E2024",
    textLight: "#FFFFFF",
    border: "#E1E4E8",
    fontFamily: "Arial",
  },
};

export const TEMPLATE_PRESETS = {
  // ── 1. Word: Academic Assignment / Research Paper ──────────────────────────
  academic_report: {
    app: "word",
    theme: "politecnico_maroon",
    pageSetup: {
      paperSize: "letter",
      orientation: "portrait",
      topMargin: 72,
      bottomMargin: 72,
      leftMargin: 72,
      rightMargin: 72,
    },
    defaultCoverPage: {
      institution: "INSTITUTO POLITÉCNICO NACIONAL",
      school: 'CENTRO DE ESTUDIOS CIENTÍFICOS Y TECNOLÓGICOS N° 8\n"NARCISO BASSOLS"',
      subject: "INGENIERÍA DE SISTEMAS Y CIBERESPACIO",
      title: "INVESTIGACIÓN Y DESARROLLO TECNOLÓGICO",
      subtitle: "Arquitectura de Software y Automatización de Ecosistemas",
      group: "3IV1",
      team: ["Lara Lozano Mauricio Leonardo"],
      date: "Octubre de 2026",
    },
    sectionsOrder: [
      "cover_page",
      "table_of_contents",
      "executive_summary",
      "methodology",
      "results_and_data",
      "conclusions",
      "bibliography",
    ],
  },

  // ── 2. Word: Executive Business Proposal ──────────────────────────────────
  executive_proposal: {
    app: "word",
    theme: "modern_blue",
    pageSetup: {
      paperSize: "letter",
      orientation: "portrait",
      topMargin: 54,
      bottomMargin: 54,
      leftMargin: 54,
      rightMargin: 54,
    },
    headerFooter: {
      headerText: "PROPUESTA DE ARQUITECTURA TÉCNICA · CONFIDENCIAL",
      footerText: "Aero Host Suite · Página ",
      differentFirstPage: true,
    },
    sectionsOrder: [
      "executive_summary",
      "value_proposition",
      "technical_solution",
      "financial_investment",
      "roadmap_and_milestones",
      "governance_and_support",
    ],
  },

  // ── 3. Excel: Financial Model & KPI Dashboard ─────────────────────────────
  financial_model: {
    app: "excel",
    theme: "modern_blue",
    sheets: [
      {
        name: "Dashboard KPI",
        headers: ["Indicador Clave (KPI)", "Meta Anual", "Real Acumulado", "Variación %", "Estado"],
        defaultData: [
          ["Ingresos Totales", 1200000, 1285000, "=(C2-B2)/B2", "Superado"],
          ["Costos Operativos", 450000, 420000, "=(C3-B3)/B3", "Eficiente"],
          ["EBITDA", "=B2-B3", "=C2-C3", "=(C4-B4)/B4", "Óptimo"],
          ["Margen Operativo", "=B4/B2", "=C4/C2", "=C5-B5", "Favorable"],
        ],
        freezePanes: "A2",
      },
      {
        name: "Detalle Mensual",
        headers: ["Mes", "Ventas Directas", "Suscripciones", "Servicios", "Total Ingresos"],
        defaultData: [
          ["Enero", 85000, 25000, 10000, "=SUM(B2:D2)"],
          ["Febrero", 92000, 27000, 12000, "=SUM(B3:D3)"],
          ["Marzo", 104000, 29000, 15000, "=SUM(B4:D4)"],
        ],
        freezePanes: "B2",
      },
    ],
  },

  // ── 4. PowerPoint: Executive Pitch Deck (16:9 Widescreen) ───────────────────
  presentation_deck: {
    app: "powerpoint",
    theme: "modern_blue",
    aspectRatio: "16:9",
    slides: [
      {
        layout: "title",
        title: "Aero Host Suite",
        body: "Automatización de Alto Rendimiento para Microsoft Office",
      },
      {
        layout: "title_content",
        title: "El Problema: Fricción en Documentos Reales",
        body: "• Formatos rotos y falta de estandarización\n• Sobrecarga de tokens con lecturas de archivos completos\n• Imposibilidad de interactuar en vivo con la pantalla del usuario",
      },
      {
        layout: "two_content",
        title: "Solución: Doble Modo LIVE y Headless",
        body: "Control en vivo vía COM nativo sin recargar ni cerrar ventanas activas.",
      },
      {
        layout: "title_content",
        title: "Métricas de Rendimiento",
        body: "• Ahorro de tokens: -88% a -99%\n• Tiempo de respuesta: < 400ms\n• Cobertura de pruebas: 100% exitosa",
      },
      {
        layout: "title_content",
        title: "Próximos Pasos y Conclusión",
        body: "Despliegue autónomo en producción con cero ceremonia de permisos.",
      },
    ],
  },
};

export function getTheme(themeName) {
  return THEMES[themeName] || THEMES.modern_blue;
}

export function getPreset(presetName) {
  return TEMPLATE_PRESETS[presetName] || null;
}
