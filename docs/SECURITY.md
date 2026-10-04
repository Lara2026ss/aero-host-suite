# FLUXER XZ · Seguridad

Este documento describe la arquitectura de seguridad, la clasificación de riesgos y los mecanismos de autorización de **FLUXER XZ - PUBLIC RELEASE V26-0.1**.

---

## 1. Filosofía de Seguridad

FLUXER XZ opera bajo el principio de **mínimo privilegio explícito** y **cero simulación**:
1. La IA no puede ejecutar acciones destructivas o de alto riesgo sin consentimiento visible del usuario.
2. Cada solicitud de autorización muestra una descripción clara, amable y precisa:
   - Qué acción desea realizar la IA.
   - Sobre qué archivo o recurso del sistema.
   - Cuál es el motivo.
   - Qué nivel de riesgo implica.
   - Qué ocurre si el usuario rechaza la acción.
3. Todas las ventanas y notificaciones portan el título oficial fijo `"Fluxer XZ · Seguridad"` y el icono oficial de escudo (`icon.ico`).

---

## 2. Niveles de Riesgo Canónicos

| Nivel de Riesgo | Descripción | Ejemplos de Acciones | Política de Autorización |
|---|---|---|---|
| **LOW** | Operaciones de solo lectura o diagnósticos no intrusivos que no modifican el sistema. | `files.read_file`, `system.get_info`, `guide.get_catalog`, `shortcuts.list_shortcuts` | Permitido automáticamente por defecto. |
| **MEDIUM** | Creación o modificación de archivos de usuario en rutas permitidas; pruebas de red locales. | `files.write_file`, `files.patch_file`, `web.download_file`, `shortcuts.create_shortcut` | Permitido una vez o por sesión. Elevación previa lo aprueba automáticamente. |
| **HIGH** | Ejecución de comandos en la terminal, inicio/detención de procesos, gestión de paquetes o servicios. | `terminal.run_command`, `system.kill_process`, `system.install_package`, `media.capture_desktop` | Requiere confirmación explícita o sesión previamente elevada. Envía notificación nativa de Windows con AUMID. |
| **DESTRUCTIVE** | Eliminación permanente de archivos, depuración de carpetas o cierre forzado crítico. | `files.delete_path` (permanente), `shortcuts.delete_shortcut` | **SIEMPRE requiere confirmación interactiva explícita**, incluso si la sesión está elevada. |

---

## 3. Opciones de Autorización en Diálogo

Cuando una acción requiere aprobación, el diálogo oficial ofrece 3 opciones claras:
- **Permitir una vez**: La acción se ejecuta una sola vez. Siguientes ejecuciones volverán a consultar.
- **Permitir esta sesión**: La acción queda autorizada en memoria para la sesión activa del servidor MCP, evitando diálogos repetitivos molestos.
- **Rechazar**: La acción se cancela inmediatamente de forma segura, informando a la IA sin alterar el sistema operativo.

---

## 4. Rutas Críticas Protegidas

Las siguientes rutas y patrones del sistema operativo están **estrictamente bloqueadas** por el router de seguridad (`core/permissions.mjs`):
- Carpetas del sistema Windows: `C:\Windows\System32`, `C:\Windows\SysWOW64`.
- Hives de registro críticas de Windows: `SAM`, `SYSTEM`, `SECURITY`, `SOFTWARE`, `DEFAULT`.
- Credenciales de autenticación: Claves SSH (`.ssh/id_rsa`, `id_ed25519`, `authorized_keys`), credenciales AWS (`.aws/credentials`), perfiles Azure (`.azure/`), configuraciones Kubernetes (`.kube/config`).
- Tablas de resolución de red local: `drivers\etc\hosts`.
- Sistemas Unix/WSL: `/etc/shadow`, `/etc/passwd`, `/etc/sudoers`.
- Intentos de Directory Traversal (`..\..\Windows\System32`).

Cualquier intento de acceso o modificación a estas rutas es interceptado inmediatamente con la decisión `blocked_system_integrity` y registrado en el log forense, sin importar si la sesión se encuentra elevada.

---

## 5. Auditoría Forense (`data/security/audit.log`)

Cada verificación, cambio de elevación o decisión de permiso se registra en formato JSON inmutable en `data/security/audit.log`, conteniendo:
- Marca temporal ISO 8601.
- Herramienta y acción solicitada.
- Recurso involucrado.
- Nivel de riesgo.
- Decisión final (`allowed_low_risk`, `allowed_elevated_session`, `granted_session`, `granted_once`, `denied_by_user`, `blocked_system_integrity`).

---

## 6. Revocación Inmediata de Elevación

La elevación puede ser revocada en cualquier instante tanto programáticamente a través de `core/permissions.mjs` (`revokeElevation()`), vía la herramienta `security.revoke_elevation` desde el cliente MCP, o pulsando el botón correspondiente en la aplicación nativa `FluxerXZ.exe`. Al revocar:
- Se eliminan todos los permisos temporales en memoria.
- Se restablece el estado base a no elevado.
- Se emite una entrada de auditoría forense con marca temporal.

