# Guía para Contribuir a Aero Fluxer X

¡Gracias por tu interés en contribuir a Aero Fluxer X!

## Principios Fundamentales
1. **Seguridad Absoluta**: NUNCA envíes credenciales, claves de API, tokens ni rutas personales.
2. **Aislamiento de Datos**: El código del repositorio debe permanecer apátrida (*stateless*). Todo dato generado debe dirigirse al almacenamiento local del usuario (`core/storage-paths.mjs`).
3. **Compatibilidad SemVer**: Cualquier cambio debe respetar el versionado SemVer 2.0 y actualizar los contratos en `contracts/fluxer_mcp_tools.json`.
4. **Verificación Adversarial**: Todas las herramientas deben probarse contra el motor `doctor.mjs` y pasar `tests/test_fluxer_suite.mjs`.

## Flujo de Trabajo
1. Crea un Fork del repositorio.
2. Crea una rama para tu feature: `git checkout -b feature/mi-mejora`.
3. Implementa tus cambios siguiendo la arquitectura modular de 10 dominios.
4. Ejecuta las pruebas:
   ```bash
   node tests/test_fluxer_suite.mjs
   node tests/test_distribution_and_updater.mjs
   npm run doctor
   ```
5. Realiza un commit con un mensaje descriptivo y abre un Pull Request.
