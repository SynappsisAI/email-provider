# Changelog

Todos los cambios relevantes de este repositorio. Formato basado en
[Keep a Changelog](https://keepachangelog.com); las versiones siguen SemVer y se
publican como tags de git (`vX.Y.Z`). Entradas en español.

## [Unreleased]

## [0.4.1] — 2026-09-23

### Fixed

- **El reply nativo de Gmail ignoraba el `Reply-To`** y respondía siempre al `From`. Ahora
  responde al `Reply-To` cuando el remitente lo define (RFC 5322), igual que el reply nativo
  de Microsoft Graph y cualquier cliente de correo — p.ej. un sistema de tickets que envía
  desde `notificaciones@` con `Reply-To: ticket-123@` recibe la respuesta en el ticket.
  Si el `Reply-To` existe pero no trae ninguna dirección utilizable, se responde al `From`.
- **Seguridad: el reply nativo de Gmail podía responder a más direcciones que las que el
  consumidor validó.** `readMessage` parsea el `From` como UNA dirección, pero el reply lo
  partía por comas: `From: evil@evil.com, <ok@allowed.com>` se validaba como `ok@allowed.com`
  y la respuesta iba también a `evil@evil.com`. Ahora el reply usa el mismo parse del `From`.
  Preexistente (≤ v0.4.0).
- **Direcciones con coma en el nombre (`"Soporte, Acme" <t@acme.com>`) se partían en dos**
  y las sin nombre (`<a@x.com>`) conservaban los `<>` en la dirección (Gmail). El parser
  ahora respeta comillas y corchetes (si quedan sin cerrar, vuelve a partir por cada coma,
  para no esconder direcciones dentro de un token malformado), y descarta entradas vacías.

### Added

- **`MessageFull.isMailingList`**: `true` si el mensaje trae `List-Id` o `List-Post` (listas
  de correo / Google Groups). Ahí el `Reply-To` suele apuntar a toda la lista, así que quien
  arma una respuesta privada puede preferir el `From`.

## [0.4.0] — 2026-09-22

### Added

- **`replyToMessage` y `forwardMessage` aceptan `attachments`** (típicamente imágenes
  inline con `contentId`, p.ej. el logo de una firma). En Microsoft Graph, con adjuntos
  se usa el flujo de borrador (`createReply`/`createReplyAll`/`createForward` → PATCH del
  cuerpo con nuestro HTML arriba del citado → adjuntos → `send`; si algo falla, el borrador
  se borra); sin adjuntos se mantiene el endpoint de un paso de siempre. En Gmail, los
  adjuntos van al MIME del reply/forward. Retrocompatible. En Graph cada adjunto de ese
  camino debe pesar < 3MB (subida simple). Los borradores de reply copian además las
  imágenes inline del original (best-effort: se omiten las ≥3MB y las que ya trae el
  borrador; un error ahí nunca bloquea el reply), para que la cita no salga rota.
- **`MessageFull.replyTo`** (header `Reply-To`). Graph entrega el reply nativo al
  `Reply-To`, no al `From`: los consumidores que validan destinatarios (allowlists) deben
  validarlo también — antes era invisible.

### Fixed

- **El forward de Gmail perdía los adjuntos del correo original.** Ahora los lleva mientras
  el mensaje codificado (base64, incluidos los adjuntos propios) quepa en ~24MB; por encima,
  sale solo el cuerpo, como antes. Una parte queda inline solo
  si el HTML la referencia con `cid:` o su disposición es `inline` (Outlook pone
  `Content-ID` también en PDFs); las imágenes inline sin nombre se conservan con un nombre
  de respaldo.
- **Nombres de adjunto con acentos o comillas rompían el header MIME.** Ahora se codifican
  con RFC 2231 (`filename*=UTF-8''…`).

## [0.3.1] — 2026-09-22

### Fixed

- **`multipart/related` ahora declara `type="text/html"`** (parámetro obligatorio según
  RFC 2387). Sin él, un parser estricto o un gateway de seguridad podía mostrar la imagen
  inline como adjunto suelto. Gmail/Outlook/Apple Mail ya lo toleraban.

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
