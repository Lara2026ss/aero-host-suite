# Configuración de Clientes MCP para Aero Host Suite

Aero Host Suite soporta una integración transparente (Zero-Friction) con múltiples clientes de Inteligencia Artificial que soportan el protocolo MCP (Model Context Protocol).

Si necesitas configurar un cliente de forma manual, utiliza las siguientes instrucciones.

## Requisitos Previos

- **Node.js**: v22.0 o superior (`node -v`)
- **Directorio del motor**: Localiza dónde se encuentra tu instalación de Aero Host Suite:
  `C:/Users/mauri/OneDrive/Documents/Systems/Flux-mcp-doc/Windows MCP/Aero Host Suite`
  Entrypoint principal: `server.js` (o `server.mjs`).

---

## 1. Google Antigravity

Antigravity es un entorno de codificación local de Google que soporta MCP nativamente.

1. Abre el archivo de configuración de Antigravity MCP:
   `C:/Users/mauri/.gemini/antigravity/mcp_config.json`
2. Añade la entrada `Aero_Host_Suite` bajo `mcpServers`:

```json
{
  "mcpServers": {
    "Aero_Host_Suite": {
      "command": "C:/Program Files/nodejs/node.exe",
      "args": [
        "C:/Users/mauri/OneDrive/Documents/Systems/Flux-mcp-doc/Windows MCP/Aero Host Suite/server.js"
      ],
      "cwd": "C:/Users/mauri/OneDrive/Documents/Systems/Flux-mcp-doc/Windows MCP/Aero Host Suite"
    }
  }
}
```

3. Guarda el archivo. Antigravity recargará los MCP automáticamente al iniciar un nuevo chat.

---

## 2. Claude Desktop (Anthropic)

Claude Desktop para Windows soporta herramientas locales vía MCP.

1. Abre el archivo de configuración:
   `%APPDATA%\Claude\claude_desktop_config.json`
2. Añade la entrada de Aero Host Suite:

```json
{
  "mcpServers": {
    "Aero_Host_Suite": {
      "command": "node",
      "args": [
        "C:/Users/mauri/OneDrive/Documents/Systems/Flux-mcp-doc/Windows MCP/Aero Host Suite/server.js"
      ]
    }
  }
}
```

3. Reinicia Claude Desktop por completo. El ícono del martillo mostrará las 10 herramientas de Aero Host Suite disponibles.

---

## 3. Cursor IDE

1. Abre Cursor Settings (`Ctrl + ,`) -> Features -> MCP Servers.
2. Añade un nuevo servidor:
   - **Name**: `Aero_Host_Suite`
   - **Type**: `command`
   - **Command**: `node "C:/Users/mauri/OneDrive/Documents/Systems/Flux-mcp-doc/Windows MCP/Aero Host Suite/server.js"`

---

## 4. Verificación de Conectividad

Para verificar que el servidor inicia y responde correctamente desde la terminal:

```powershell
node server.mjs --check
# o ejecutar la suite de pruebas:
npm test
```
