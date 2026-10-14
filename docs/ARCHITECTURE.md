# Aero Host Suite — Arquitectura del Sistema (V26-0.2)

Reconstrucción modular y optimizada. Aero Host Suite expone exactamente **10 dominios consolidados** al modelo de IA, maximizando la eficiencia de tokens y eliminando la sobrecarga cognitiva.

Cada dominio funciona como un router unificado que resuelve internamente acciones de automatización del sistema host en Windows 11.

---

## 1. Los 10 Dominios Oficiales

1. **files** (`tools/files.mjs`) — Lectura por rango, parches quirúrgicos por texto ancla (`patch_file`), escritura incremental, búsqueda recursiva, metadatos, copia y borrado seguro.
2. **terminal** (`tools/terminal.mjs`) — Ejecución síncrona de comandos PowerShell/CMD, procesos en segundo plano con control de trabajos (`jobs`) y scripts temporales.
3. **system** (`tools/system.mjs`) — Telemetría de hardware (CPU, RAM, batería), gestión de procesos por PID, red (ping, DNS, puertos) y gestión de paquetes del sistema (winget, npm).
4. **web** (`tools/web.mjs`) — Búsqueda web DuckDuckGo directa, extracción limpia de markdown desde páginas web y descargas seguras con verificación SHA-256.
5. **documents** (`tools/documents.mjs`) — Automatización en vivo (COM) y headless de Word, Excel, PowerPoint y PDF (estructuras, textos, tablas, diapositivas).
6. **media** (`tools/media.mjs`) — Captura visual de pantalla completa y ventanas activas, y gestión de colas de impresión de Windows.
7. **shortcuts** (`tools/shortcuts.mjs`) — Orquestación de macros multi-paso, gestión de plantillas Mustache, validación y ejecución atómica con rollback.
8. **upd** (`tools/upd.mjs`) — Interfaz MCP para consulta, descarga, aplicación y reversión de actualizaciones remotas.
9. **feedback** (`tools/feedback.mjs`) — Registro de retroalimentación técnica, anonimización automática de rutas de usuario (`~`) y generador de issues en GitHub.
10. **guide** (`tools/guide.mjs`) — Catálogo interactivo bajo demanda, recetas de maquetación de documentos y estado del servidor.

---

## 2. Flujo de Ejecución de una Llamada

```
Cliente MCP (Antigravity / Claude Desktop / Cursor)
   → ListTools (10 tools oficiales)
   → CallTool { tool: "documents", action: "create", args: { app: "word" } }
      → Router.execute(tool, action, args)
         → permissions.checkPermission() (validación de riesgo sin fricción)
         → registry.resolve(tool, action) → invocación de handler real
         → compact(response)              (reducción real de tokens 88%-99%)
         → logger.log()                   (auditoría local)
      ← resultado compacto con handles (f#1, doc#1, stash#id)
```

---

## 3. Infraestructura del Núcleo (`core/`)

- `core/runtime.mjs`: Inicialización del entorno, gestión de rutas de datos (`data/`) y coordinación de subsistemas.
- `core/registry.mjs`: Registro dinámico y lazy-loading de módulos de herramientas.
- `core/router.mjs`: Despachador de acciones con manejo estructurado de errores.
- `core/compact.mjs`: Motor de compactación y optimización de tokens (Content Stashes, Handles cortos, Diffing quirúrgico).
- `core/version.mjs`: Motor SSOT de versionado Semver y nombres de producto canónicos.
- `core/logger.mjs`: Logger estructurado en `data/logs/aero-host.log`.
- `updater/updater.mjs`: Motor de actualización autónoma con verificación de hash SHA-256 y rollback automático.
