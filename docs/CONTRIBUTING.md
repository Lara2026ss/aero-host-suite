# Guía para Contribuir a Aero Host Suite

¡Gracias por tu interés en contribuir a Aero Host Suite!

## Principios Fundamentales
1. **Seguridad Absoluta**: NUNCA envíes credenciales, claves de API, tokens ni rutas personales.
2. **Aislamiento de Datos**: El código del repositorio debe permanecer apátrida (*stateless*). Todo dato generado debe dirigirse al almacenamiento local del usuario (`data/`).
3. **Compatibilidad SemVer**: Cualquier cambio debe respetar el versionado SemVer 2.0 y el SSOT en `VERSION`.
4. **Verificación Adversarial**: Todas las herramientas deben probarse físicamente y pasar `node tests/run_all_tests.mjs`.

## Flujo de Trabajo
1. Crea un Fork del repositorio.
2. Crea una rama para tu feature: `git checkout -b feature/mi-mejora`.
3. Implementa tus cambios siguiendo la arquitectura modular de 10 dominios.
4. Ejecuta las pruebas:
   ```powershell
   npm test
   # o:
   node tests/run_all_tests.mjs
   ```
5. Realiza un commit con un mensaje descriptivo y abre un Pull Request.
