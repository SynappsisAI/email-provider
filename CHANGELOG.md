# Changelog

Todos los cambios relevantes de este repositorio. Formato basado en
[Keep a Changelog](https://keepachangelog.com); las versiones siguen SemVer y se
publican como tags de git (`vX.Y.Z`). Entradas en español.

## [Unreleased]

## [0.1.0] — 2026-04-12

### Added

- **Librería de correo agnóstica de proveedor.** Envuelve Google Workspace (Gmail
  API) y Microsoft 365 (Graph API) tras una única interfaz `EmailProvider`, para
  que los servicios de Synappsis envíen, listen, lean, busquen, respondan,
  reenvíen y gestionen adjuntos sin depender del backend de cada buzón. Mecánica
  pura: sin servidor HTTP, sin autorización, sin base64 en la API pública (los
  adjuntos entran y salen como `Buffer`).
