# FLUXER MCP 5.0 — Arquitectura

Reconstrucción completa. FLUXER expone exactamente **15 dominios** al modelo de IA.
Cada dominio es un router que resuelve internamente docenas de subacciones reales.

## Dominios visibles

1. **files** — lectura, escritura, búsqueda, directorios, comprensión/extracción de archivos
2. **system** — CPU, GPU, RAM, storage, batería, temperatura, procesos, servicios
3. **memory** — memoria persistente (SQLite) y compresión de contexto
4. **development** — proyectos, dependencias (npm/pip/cargo), git, formateo/lint
5. **web** — búsqueda, descarga, extracción de texto/links/imágenes
6. **network** — interfaces, ping, traceroute, DNS, port scan
7. **automation** — tareas, workflows, comandos en segundo plano
8. **gaming** — Steam/Heroic/Lutris/Proton/Roblox/Minecraft, optimización
9. **terminal** — ejecución de comandos, scripts, sesiones, kill
10. **database** — sqlite/postgres/mysql vía CLI, backup/restore
11. **media** — imágenes (imagemagick), audio/video (ffmpeg)
12. **browser** — apertura de páginas, screenshots, extracción HTML
13. **security** — hashes, verificación, permisos, auditoría
14. **packages** — dnf/flatpak: instalar, actualizar, repositorios
15. **process** — listar, iniciar, detener, suspender, matar procesos

## Flujo de una llamada

```
Cliente MCP (Claude Desktop)
   → ListTools (15 tools)
   → CallTool { tool: "system", action: "get_ram_info", args: {} }
      → Router.execute()
         → permissions.assertAllowed()   (siempre permitido en este runtime)
         → circuitBreaker.assert()
         → registry.resolve(tool, action) → handler real
         → handler(args, runtime)
         → compact(response)              (reduce tokens)
         → memory.recordCall()            (auditoría/historial)
      ← respuesta compactada
```

## Infraestructura común (`core/`)

Cada dominio comparte la misma infraestructura de `runtime.mjs`:

- `runtime.run()` — ejecución de comandos shell con timeout, cola de concurrencia
- `runtime.memory` — SQLite (historial, permisos, conocimiento, notas)
- `runtime.permissions` — niveles guest→developer, grants temporales
- `runtime.metrics` — snapshot de métricas de proceso
- `runtime.circuitBreaker` — corta rutas que fallan repetidamente
- `runtime.logger` — logs estructurados JSON rotados
- `core/compact.mjs` — reduce el tamaño de las respuestas para modelos pequeños

## Por qué 15 y no 100+

Un modelo de 3B–8B razona peor con 100+ herramientas visibles. Compactando a
15 dominios de alto nivel, el modelo solo necesita decidir _qué dominio_ y
_qué acción_, y el router interno se encarga de resolver, validar, cachear,
ejecutar y registrar la llamada real.
