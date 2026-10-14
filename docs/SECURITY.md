# Aero Host Suite — Política de Seguridad y Control de Host (V26-0.2)

Este documento describe la arquitectura de seguridad, la clasificación de riesgos y los salvaguardas operacionales de **Aero Host Suite (V26-0.2)**.

---

## 1. Filosofía de Seguridad y Salvaguardas Mínimas

Aero Host Suite opera bajo el principio de **ejecución deliberada**, **transparencia total** y **cero simulación**:
1. El servidor se ejecuta con los privilegios del proceso host de Node.js en Windows.
2. No se ejecutan comandos destructivos fuera del alcance del usuario sin supervisión.
3. Se mantiene una distinción clara entre eliminar subsistemas de autorización intrusivos/bloqueantes y preservar las salvaguardas fundamentales requeridas para la integridad del host.

---

## 2. Clasificación de Riesgos Operacionales

| Nivel de Riesgo | Descripción | Herramientas Típicas | Salvaguardas |
|---|---|---|---|
| **LOW** | Lectura de archivos, catálogos y diagnósticos no intrusivos. | `files.read_file`, `system.get_info`, `guide.get_catalog`, `shortcuts.list_shortcuts` | Ejecución directa e inmediata. |
| **MEDIUM** | Creación y modificación de archivos de proyecto, descargas y edición de documentos Office. | `files.write_file`, `files.patch_file`, `documents.insert_section`, `shortcuts.create_shortcut` | Operación quirúrgica, creación automática de directorios padre y preservación de codificación UTF-8. |
| **HIGH** | Ejecución de comandos en la terminal, control de procesos o captura de pantalla. | `terminal.run_command`, `terminal.run_background`, `system.kill_process`, `media.capture_screen` | Aislamiento de shell, retorno estructurado de stdout/stderr y control de PIDs. |
| **DESTRUCTIVE** | Eliminación de archivos o directorios. | `files.delete_path` | Reversible mediante papelera cuando sea posible; confirmación explícita para borrados recursivos. |

---

## 3. Integridad del Sistema Operativo

Aero Host Suite previene el acceso o modificación accidental a rutas críticas del sistema Windows:
- Carpetas del sistema: `System32`, `SysWOW64`.
- Archivos de configuración de credenciales del host (`.ssh`, credenciales cloud).
- Archivo hosts de red local.

---

## 4. Registro y Auditoría Local

Todas las operaciones se auditan de forma estructurada en `data/logs/aero-host.log` para trazabilidad local completa, sin enviar telemetría a servidores de terceros.
