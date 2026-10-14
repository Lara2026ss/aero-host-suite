# Aero Host Suite — Official Release V26-0.2

<div align="center">
  <br/>
  <strong>Servidor MCP de Alto Rendimiento para Windows con Modo Compacto Real, 10 Herramientas Nativas y Automatización Integral</strong>
  <br/>
  <em>Versión Oficial: V26-0.2 · Arquitectura Limpia Modular</em>
</div>

---

## 🌟 Características Principales

- **Modo Compacto de Próxima Generación (`core/compact.mjs`)**:
  - Ahorro de tokens real y medido del **88% al 99%** sin trucos contraproducentes (cero base64, cero morse, cero binarios).
  - **Content Stashes (`stash#id`)**: La IA envía bloques largos una sola vez y los reutiliza por referencia en múltiples archivos u operaciones.
  - **Handles Cortos (`f#1`, `doc#1`)**: Operaciones sobre archivos y documentos sin repetir rutas extensas en cada mensaje.
  - **Escritura Incremental (`patch_file`)**: Parches quirúrgicos por texto ancla o rango de líneas en vez de reescribir archivos enteros.
  - **Lectura Inteligente y Caché Local**: Detección de "sin cambios" mediante hash SHA-256 local y lecturas por rango de líneas.
- **Automatización Nativa de Documentos Office (`tools/documents.mjs`)**:
  - Control en vivo y headless vía COM nativo para Word, Excel y PowerPoint.
  - Generación de portadas institucionales, tablas con formato, estilos tipográficos jerárquicos y exportación a PDF.
- **Motor de Atajos y Macros Multi-Paso (`tools/shortcuts.mjs`)**:
  - Ejecución de pipelines complejos con una sola llamada.
  - Soporte de plantillas con variables `{{param}}`, dry-run y rollback automático ante fallos.
  - Atajos de producción preconfigurados en `data/shortcuts/`.
- **Actualizador Autónomo Zero-Dependency (`updater/updater.mjs`)**:
  - Actualización directa con verificación criptográfica SHA-256.
  - Reemplazo atómico con respaldo previo en `data/backups/`.
  - Verificación física de versión en disco tras actualización y reversión automática.
- **Puente de Feedback Privado (`tools/feedback.mjs`)**:
  - Almacenamiento local en `data/feedback/` con anonimización estricta de rutas personales (`~`).
  - Generador de enlaces directos a GitHub Issues sin exponer credenciales ni tokens.

---

## 🛠️ Catálogo de las 10 Herramientas Oficiales

| Herramienta | Archivo | Descripción y Capacidades |
|---|---|---|
| **files** | `tools/files.mjs` | Lectura por rango, parches quirúrgicos por ancla, escritura incremental, búsqueda recursiva, copia y borrado seguro. |
| **terminal** | `tools/terminal.mjs` | Ejecución de comandos powershell/cmd de forma síncrona, procesos en segundo plano con control de trabajos y scripts temporales. |
| **system** | `tools/system.mjs` | Diagnóstico de hardware (CPU, RAM, batería), gestión de procesos por PID, red (ping, DNS, puertos) y paquetes (winget, npm). |
| **web** | `tools/web.mjs` | Búsqueda web DuckDuckGo sin intermediarios, extracción de texto en markdown y descargas seguras con hash. |
| **documents** | `tools/documents.mjs` | Automatización en vivo (COM) y headless de Word, Excel, PowerPoint y PDF (estructuras, textos, tablas, diapositivas). |
| **media** | `tools/media.mjs` | Captura visual no intrusiva de pantalla completa y ventanas activas, y control de colas de impresión de Windows. |
| **shortcuts** | `tools/shortcuts.mjs` | Orquestación de macros multi-paso, gestión de plantillas Mustache, validación y ejecución con rollback. |
| **upd** | `tools/upd.mjs` | Interfaz MCP para consulta, descarga, aplicación y reversión de actualizaciones remotas. |
| **feedback** | `tools/feedback.mjs` | Registro de reportes y sugerencias locales, exportación y apertura asistida de issues en GitHub. |
| **guide** | `tools/guide.mjs` | Catálogo dinámico bajo demanda, documentación perezosa de herramientas y estado del servidor. |

---

## 📊 Medición Real del Ahorro de Tokens (Compact Mode)

Mediciones físicas reales ejecutadas sobre la suite de pruebas:

| Escenario | Línea Base Anterior | Aero Host Suite V26-0.2 | Reducción Real | Método Empleado |
|---|---|---|---|---|
| **Catálogo de Herramientas** | 19,999 tokens | 2,373 tokens | **-88.1%** | Consolidación a 10 herramientas y esquemas compactos |
| **Lectura de Código** | 36,827 tokens | 257 tokens | **-99.3%** | Lectura por rango (`startLine`/`endLine`) y handles |
| **Re-lectura Sin Cambios** | 36,827 tokens | 24 tokens | **-99.9%** | Verificación local de hash SHA-256 |
| **Modificación de 500 Líneas** | 13,949 tokens | 66 tokens | **-99.5%** | `patch_file` quirúrgico por texto ancla |
| **Reutilización de Código** | 13,949 tokens | 17 tokens | **-99.9%** | `stash#<id>` por referencia |
| **Edición de Documento Office** | 3,356 tokens | 30 tokens | **-99.1%** | Operaciones estructurales (`insert_paragraph`) y handles |

---

## 🚀 Conexión al Cliente MCP

Agregue Aero Host Suite a su archivo de configuración MCP (`mcp_config.json` en Antigravity, Claude Desktop o Cursor):

```json
{
  "mcpServers": {
    "Aero_Host_Suite": {
      "command": "node",
      "args": [
        "C:/Users/mauri/OneDrive/Documents/Systems/Flux-mcp-doc/Windows MCP/Aero Host Suite/server.mjs"
      ]
    }
  }
}
```

O utilice el punto de entrada retrocompatible `server.js`.

---

## 🧪 Pruebas Físicas

Para ejecutar la suite completa de pruebas:

```powershell
npm test
# o directamente:
node tests/run_all_tests.mjs
```

---

## 📜 Licencia

MIT License — Copyright (c) 2026 Aero Core Team
