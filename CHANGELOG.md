# Changelog

Todos los cambios relevantes de este repositorio. Formato basado en
[Keep a Changelog](https://keepachangelog.com); las versiones siguen SemVer y se
publican como tags de git (`vX.Y.Z`). Entradas en español.

## [Unreleased]

## [0.3.0] — 2026-09-22

### Added

- **Imágenes inline (CID) en `sendEmail`.** `Attachment` acepta un `contentId`
  opcional: el archivo se incrusta en el cuerpo HTML (referenciado como
  `<img src="cid:<contentId>">`) en vez de ir como adjunto descargable. Gmail lo
  arma como `multipart/related` (dentro de `multipart/mixed` si además hay adjuntos
  normales); Microsoft Graph usa `isInline` + `contentId`. Habilita firmas de correo
  con logo (caso: firma corporativa del agente de Matelpa en AgentFleet).
  Retrocompatible: sin `contentId` todo se comporta igual que antes.

## [0.2.0] — 2026-07-17

### Changed

- **Reemplazado el meta-paquete `googleapis` (~125MB) por el paquete scoped
  `@googleapis/gmail` (solo Gmail).** El provider de Google usaba únicamente la API
  de Gmail pero arrastraba TODAS las APIs de Google. El cambio es mecánico y no toca
  el contrato público `EmailProvider`: mismos tipos `gmail_v1`, misma factory
  `gmail({version,auth})`, y `GoogleAuth` se toma del re-export `auth` del propio
  `@googleapis/gmail` (sin dependencia directa extra, evita choque de versiones de
  `google-auth-library`). Reduce el footprint instalado y, aguas abajo, la latencia
  de cold start del worker Lambda de AgentFleet, donde esta librería es git-dependency.
  Cierra #2. **Follow-up en AgentFleet:** bumpear el pin git de
  `@synappsis/email-provider` para recoger el cambio.

### Added

- **Librería de correo agnóstica de proveedor.** Envuelve Google Workspace (Gmail
  API) y Microsoft 365 (Graph API) tras una única interfaz `EmailProvider`, para
  que los servicios de Synappsis envíen, listen, lean, busquen, respondan,
  reenvíen y gestionen adjuntos sin depender del backend de cada buzón. Mecánica
  pura: sin servidor HTTP, sin autorización, sin base64 en la API pública (los
  adjuntos entran y salen como `Buffer`).
