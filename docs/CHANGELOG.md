# CHANGELOG — FLUXER XZ

All notable changes to FLUXER XZ will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and follows the product versioning standard `V{YEAR}-{VERSION}` (internally semver mapped).

---

## [V26-0.1] — 2026-10-04

### Added
- **Single Source of Truth Version Engine**: Centralized in root `VERSION` file with bidirectional semver mapping (`V26-0.1` <-> `26.0.1`) and automated version comparison.
- **Redesigned Compact Mode (`core/compact.mjs`)**:
  - Content Stashes (`stash#<id>`) for sending large payloads once and referencing across steps.
  - Short Handles (`f#<id>`, `doc#<id>`) avoiding repeated long paths in conversations.
  - Range Reading & Surgical Diffs (`patch_file` via anchor text or line range).
  - Local hash caching returning "Sin cambios" when files are unmodified.
  - Measured 88% to 99% token reduction across typical coding and document operations without binary/morse/base64 compression tricks.
- **Autonomous Remote Updater (`updater/updater.mjs`)**:
  - Zero external dependencies (no Git CLI required; pure Node.js native engine).
  - SHA-256 integrity verification before archive extraction.
  - Automatic backup to `data/backups/` and physical disk verification after file replacement.
  - Automatic rollback on copy failure or version mismatch.
- **Unified Security & Brand Identity (`core/permissions.mjs`, `core/notify.mjs`)**:
  - Fixed title "Fluxer XZ · Seguridad" with official minimalist shield insignia.
  - Windows native toast notifications with registered `AppUserModelID` (`FluxerXZ.Security.V26`) and `appLogoOverride`.
  - 4 canonical risk levels: `LOW`, `MEDIUM`, `HIGH`, `DESTRUCTIVE`.
  - Persistent session elevation respecting permissions across all tools without repetitive prompts.
  - Destructive actions always require explicit user confirmation.
- **Next-Gen Shortcuts Engine (`tools/shortcuts.mjs`)**:
  - Multi-step macro execution with Mustache-style variable interpolation (`{{param}}`).
  - Dry-run simulation and automated rollback on failure.
  - Pre-seeded with 6 production-ready starter templates.
  - Dynamic user-chosen shortcuts storage directory.
- **Privacy-Safe Feedback Bridge (`tools/feedback.mjs`)**:
  - Local feedback storage in `data/feedback/` with automated user path sanitization (`~`).
  - Pre-filled GitHub Issue URL generator requiring zero embedded tokens or webhooks.
- **Universal USB & Local Launcher**:
  - Agnostic installer detecting Node.js LTS with winget/msi assistance.
  - Start Menu shortcut creation with AUMID and high-resolution icons.

### Changed
- Reconstructed entire toolchain into exactly 11 streamlined tools: `files`, `terminal`, `system`, `web`, `documents`, `media`, `shortcuts`, `security`, `upd`, `feedback`, `guide`.
- Merged `screenshot` and `printcenter` into `media.mjs`.
- Consolidated `network` and `packages` into `system.mjs`.

### Removed
- Eliminated `doctor.mjs` and unneeded diagnostic checks.
- Purged legacy v8, v10, v11, v20, v30 and v31 codebases, relics, and temporary scripts.
