# 🪟 AERON FLUXER X MCP v9.0 — ARQUITECTURA EXCLUSIVA PARA WINDOWS 10/11

## 1. Directiva y Política de Plataforma Única

**Aeron Fluxer X MCP v9.0** es un motor MCP diseñado y optimizado **estricta y exclusivamente para Windows 10 y Windows 11 (plataforma `win32`)**.

No implementa compatibilidad ficticia con Linux ni macOS. Cualquier intento de inicialización en una plataforma distinta de `win32` es interceptado en el bootstrap (`assertWindows()`) y abortado de forma inmediata con un mensaje fatal estructurado.

---

## 2. Diagnóstico y Resolución de los 4 Fallos de Claude

### 1. Shell Asumido Incorrectamente (Bug 1)
- **Causa Raíz:** Los comandos se ejecutaban en PowerShell sin declarar el shell real y aceptaban sintaxis bash como `&&` o `||`, provocando errores crípticos del parser nativo de PowerShell.
- **Solución Aplicada:**
  - Contrato explícito de shell: `powershell` (por defecto), `cmd`, `direct`.
  - Detección previa y rechazo controlado de sintaxis bash (`&&`, `||`) con mensajes educativos que indican cómo realizar la misma operación en PowerShell (`;`, `if ($?) { ... }`, `if ($LASTEXITCODE -eq 0) { ... }`) o sugieren activar `shell: 'cmd'`.
  - Retorno de metadatos de trazabilidad: `effectiveShell`, `resolvedCommand`, `exitCode`, `durationMs`, `encoding`.

### 2. Inconsistencia de PATH y Toolchain (Bug 2)
- **Causa Raíz:** Cada módulo (`diagnostics`, `terminal`, `packages`) descubría binarios o leía variables de entorno de forma ad-hoc.
- **Solución Aplicada:**
  - Creación del módulo centralizado `core/toolchain.mjs`.
  - Normalización canónica del `PATH` (deduplicación case-insensitive, inyección de directorios críticos de Windows como `System32` y `WindowsPowerShell`).
  - Única fuente de verdad mediante `getToolchainSnapshot()`.
  - Nueva subherramienta `diagnostics.resolve_toolchain` y contrato enriquecido en `diagnostics.health_check`.

### 3. CWD No Respetado o Ignorado (Bug 3)
- **Causa Raíz:** Rutas relativas o inexistentes pasadas en `cwd` se interpretaban laxamente.
- **Solución Aplicada:**
  - Validación de existencia estricta con `existsSync` antes de cualquier spawn.
  - Resolución canónica a ruta absoluta de Windows (`path.resolve`).
  - Retorno explícito de `effectiveCwd` en la respuesta estructurada.
  - Nueva subherramienta `files.validate_workspace` para validación de límites del espacio de trabajo.

### 4. Mojibake y Corrupción de Caracteres en stdout/stderr (Bug 4)
- **Causa Raíz:** Discrepancias entre las páginas de códigos de Windows (CP850 / CP1252) y UTF-8.
- **Solución Aplicada:**
  - Inyección obligatoria de configuración UTF-8 en PowerShell:
    `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8;`
  - En CMD: `chcp 65001 >nul && ...`
  - Decodificación uniforme de buffers en UTF-8 y sanitización limpia de flujos CLIXML.

---

## 3. Matriz de Pruebas de Calidad

| Suite | Nombre | Casos Evaluados | Estado |
|---|---|---|---|
| **Suite 1** | Smoke de Plataforma | Arranque en Windows, rechazo en no-Windows | 🟢 3/3 PASS |
| **Suite 2** | Shell y Parsing | PowerShell simple/compuesto, rechazo de `&&`, modo CMD | 🟢 6/6 PASS |
| **Suite 3** | PATH & Toolchain | Deduplicación de PATH, sincronía Diagnostics-Terminal-Packages | 🟢 5/5 PASS |
| **Suite 4** | CWD Validation | CWD absoluto, relativo, rechazo de inexistente, `validate_workspace` | 🟢 4/4 PASS |
| **Suite 5** | UTF-8 & Mojibake | Acentos en español, caracteres especiales, integridad de lectura/escritura | 🟢 3/3 PASS |
| **Suite 6** | Auditoría Completa | 164 acciones a través de los 10 dominios modulares | 🟢 164/164 PASS |
| **Manual** | Protocolo Claude | 7 pasos manuales con evidencia registrada en vivo | 🟢 7/7 PASS |
