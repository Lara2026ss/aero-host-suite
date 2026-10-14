---
name: aero-host-mcp
description: Convenciones y patrones del proyecto Aero Host Suite MCP. Usar siempre que se edite, depure o extienda el MCP server Aero Host Suite en Windows MCP. Cubre arquitectura de 10 herramientas, modo compacto, automatización COM de Office, macros de shortcuts y validación rigurosa de pruebas.
---

# Aero Host Suite MCP — Convenciones del Proyecto

## Contexto del Proyecto

- **Proyecto**: Aero Host Suite (V26-0.2)
- **Rol**: Servidor MCP de Alto Rendimiento para el entorno host Windows 11.
- **Herramientas (10)**: `files`, `terminal`, `system`, `web`, `documents`, `media`, `shortcuts`, `upd`, `feedback`, `guide`.
- **Arquitectura**: Modular en Node.js ESM nativo con integración COM para Office.

## Reglas de Desarrollo

1. **Sin dependencias pesadas ni launchers obsoletos**:
   - Todo el servidor corre directamente como un proceso Node.js stdio MCP.
   - Sin notificaciones emergentes invasivas ni diálogos bloqueantes.
2. **Validación estricta de argumentos**:
   - Validar argumentos requeridos antes de tocar el filesystem o ejecutar procesos.
   - Para PIDs, validar que sea un entero positivo.
3. **Manejo de rutas en Windows**:
   - Soporte de rutas con espacios, barras `/` o `\\`.
   - Manejo de `EXDEV` para operaciones entre diferentes unidades de disco.
4. **Verificación**:
   - Ejecutar siempre `npm test` (`node tests/run_all_tests.mjs`) tras cualquier modificación.
