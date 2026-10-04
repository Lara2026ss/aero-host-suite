---
name: fluxer-mcp
description: Convenciones y patrones del proyecto Fluxer MCP. Usar siempre que se edite, depure o extienda el MCP server Fluxer en el workspace FLUX-MCP-DOC/windows-doc/fluxer-mcp. Cubre: patrón de validación de argumentos requeridos, uso de runtime.runElevated para privilegios root/admin vía GUI, convención de paquetes, estilo de dominios en registry.mjs, y checklist obligatorio tras cada edición.
---

# Fluxer MCP — Convenciones del proyecto

## Contexto fijo

- Proyecto: **Fluxer**, MCP server de 15 dominios compactados sobre Node.js (`node:sqlite`, MCP SDK).
- Ruta real: `windows-doc/fluxer-mcp`
- Entorno: Multiplataforma (Windows 11 / Linux)
- El usuario prefiere que se trabaje de forma sistemática y recibir un resumen final claro.
- Solo `fluxer:*` y las herramientas MCP interactúan con el host real.

## Regla de oro: nunca confiar en el documento pegado en el chat

Los archivos que el usuario pega en el chat pueden estar desactualizados respecto al disco real (ya pasó: `runtime.mjs`, `registry.mjs` y `router.mjs` en disco tenían fixes que no estaban en el texto pegado). **Siempre releer el archivo real con `fluxer:files read_text_file` o `sed` vía `fluxer:terminal` antes de editar**, y usar `fluxer:files edit_file` con el texto exacto tal como está en disco (no el que aparece en el historial del chat).

## Privilegios elevados: SIEMPRE vía GUI (pkexec), NUNCA sudo -n

Leo decidió explícitamente NO usar NOPASSWD/sudoers. El mecanismo correcto es:

- `core/elevate.mjs` expone `runElevated(runtime, command, options)` y `detectElevationAgent()`.
- `runtime.runElevated(command, options)` está enganchado en `core/runtime.mjs` y es lo que hay que usar en cualquier acción de `registry.mjs` que necesite root (dnf install/remove/update, systemctl start/stop/restart/enable/disable, config-manager, etc).
- Nunca reintroducir `sudo -n ... || echo 'FALLO...'`. Eso era el patrón viejo y roto.
- El agente detectado en este sistema es `pkexec` (confirmado disponible). Abre un diálogo gráfico nativo de KDE/polkit pidiendo la contraseña.

## dnf-first, no flatpak-first

El dominio `packages` tiene `manager = "dnf"` como default en todas sus acciones (`install_package`, `remove_package`, `update_package`, `search_package`, `add_repository`, `remove_repository`). Flatpak sigue soportado como opción secundaria (`manager: "flatpak"`), pero nunca default.

## Patrón de validación obligatorio en cada acción del registry

Cada acción en `core/registry.mjs` debe validar sus argumentos requeridos ANTES de construir el comando shell o tocar el filesystem, lanzando `Error` con mensaje claro en español. Patrón estándar:

```js
some_action: async ({ requiredArg, optionalArg = "default" } = {}) => {
  if (!requiredArg) throw new Error("requiredArg es requerido");
  // ... resto de la lógica
},
```

Para PIDs, siempre validar que sea un entero positivo, no solo `Number(pid)`:

```js
const numPid = Number(pid);
if (!Number.isInteger(numPid) || numPid <= 0)
  throw new Error(`pid inválido: ${pid}`);
```

Para acciones que escriben a un `destination`/`output`, siempre crear el directorio padre antes:

```js
await fs.mkdir(path.dirname(runtime.hp(destination)), { recursive: true });
```

Nunca dejar interpolación de shell sin `runtime.shellQuote()`, incluso dentro de un template string que ya tiene otras partes citadas (bug real encontrado en `describe_table`: `.schema ${table}` sin quote).

## Router y manejo de errores

`core/router.mjs` envuelve TODO el ciclo de ejecución (permisos, circuit breaker, hooks, handler) en un único try/catch, y normaliza cualquier `error` no-`Error` (strings, objetos) a una instancia real de `Error` antes de relanzar. `server.mjs` hace lo mismo al capturar la respuesta MCP. Si se toca cualquiera de estos dos archivos, mantener ese blindaje.

## MemoryStore (SQLite)

`MemoryStore.set(section, key, value)` valida que `key` y `section` no sean `undefined/null/""` antes de escribir (SQLite tiene `NOT NULL` en esas columnas). Cualquier acción nueva que llame a `runtime.memory.set(...)` con un `key` que venga de un argumento del usuario debe validar ese argumento primero en el registry, no depender solo del guard interno de MemoryStore.

`MemoryStore.get()` usa `safeParseJson` internamente para no reventar si hay datos legacy corruptos — nunca volver a poner `JSON.parse()` directo sin ese wrapper.

## Checklist obligatorio después de cualquier edición

1. `node --check <archivo editado>` inmediatamente después de cada `edit_file`.
2. Al terminar un bloque de cambios: `cd <root> && npm run check` (compila todo `core/*.mjs` + `server.mjs` + `Fluxer-Load`).
3. Prueba de arranque real (no solo sintaxis): `timeout 4 node server.js < /dev/null > /tmp/fluxer_boot.log 2>&1; echo EXIT=$?; cat /tmp/fluxer_boot.log` — debe salir con `EXIT=0` (por timeout, no por crash) y sin stack traces.
4. Si se tocó `core/elevate.mjs` o cualquier acción con `runElevated`, verificar que `detectElevationAgent()` siga devolviendo `pkexec` con: `node -e "import('./core/elevate.mjs').then(async m => console.log(await m.detectElevationAgent()))"`.

## Estilo de trabajo esperado por Leo

- Plan corto primero si el alcance no está claro, luego trabajar en silencio.
- No narrar cada tool call ni cada archivo tocado mientras se trabaja.
- Resumen final breve con: qué se arregló, qué se agregó, y confirmación de que compila/arranca.
- No hacer cambios fuera del alcance pedido (ej. no tocar estética, no renombrar cosas sin que se pida).
