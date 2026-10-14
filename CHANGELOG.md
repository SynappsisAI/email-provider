# Changelog

Todos los cambios relevantes de este repositorio. Formato basado en
[Keep a Changelog](https://keepachangelog.com); las versiones siguen SemVer y se
publican como tags de git (`vX.Y.Z`). Entradas en español.

## [Unreleased]

## [0.7.0] — 2026-10-13

### Added

- **`listSince`: mensajes desde un watermark, del más viejo al más nuevo, con token de
  continuación estable.** Devuelve los mensajes con `receivedAt >= since` en orden
  ascendente y un `nextPageToken` que retoma exactamente donde quedó la página anterior,
  aunque entre llamadas se marquen mensajes como leídos, se muevan o se borren. Corrige de
  raíz el bug de correo perdido de AgentFleet (#275): el detector de email-watch listaba
  los 50 más nuevos por tick y, en una ráfaga de más de 50, los más viejos quedaban debajo
  del watermark y nunca se procesaban. El token es un cursor por tiempo + ids ya devueltos,
  no un offset: el `@odata.nextLink` de Graph para mensajes usa `$skip` (documentado), y en
  Gmail `maxResults` tiene tope de 500 y el orden no está garantizado, así que paginar por
  offset pierde mensajes. En Gmail cada página acota una ventana de tiempo, pide metadata con
  máximo 10 llamadas simultáneas y filtra `internalDate` exacto. Una página corta o vacía
  puede traer token; solo la falta de token significa que no hay más. Cierra #8.
  **Follow-up en AgentFleet:** bumpear el pin a `v0.7.0` y usar `listSince` en el detector.

### Changed

- **BREAKING — `EmailProvider` agrega el método `listSince`.** Quien implemente la interfaz
  (mocks en tests) tiene que agregarlo; a quien solo consume la librería no le cambia nada.

## [0.6.0] — 2026-10-13

### Added

- **Campos de hilo en los mensajes.** `MessageSummary` y `MessageFull` traen ahora `cc`,
  `conversationId` (Graph `conversationId`, Gmail `threadId`), `internetMessageId` y, en
  Microsoft, `parentFolderId`. `MessageFull` además trae `uniqueBodyHtml`/`uniqueBodyText`
  (solo Microsoft: solo lo que el mensaje agregó, sin citas) y `kind` (`message`, `meeting`,
  `auto-reply`, `read-receipt`, `ndr`), calculado a partir de headers y, en Microsoft, de la
  clase de mensaje de Exchange. A diferencia de lo que sugería #7, `X-Auto-Response-Suppress`
  no cuenta como auto-respuesta: Exchange lo agrega al correo normal que llega por listas de
  distribución, y tomarlo como señal descartaría mensajes reales. En Gmail, un `.ics`
  adjunto no convierte el mensaje en `meeting`; solo la parte de invitación. Habilita que AgentFleet trate el buzón del agente como
  conversaciones (AgentFleet#275, P0 del plan `PLAN-EMAIL-AS-CHAT.md`).
- **Header de origen configurable.** La config de cualquier provider acepta `originHeader`
  (p. ej. `X-AgentFleet-Origin`). `origin` en `sendEmail`, `replyToMessage`,
  `forwardMessage` y `sendConversationMessage` lo estampa al crear el mensaje, y
  `MessageFull.origin` lo devuelve al leer. El nombre del header no queda fijo en la
  librería, que sigue siendo agnóstica del consumidor. No es una señal de seguridad:
  cualquiera puede poner un header X-.
- **`sendConversationMessage`: envío por MIME con hilo controlado por quien llama.**
  `In-Reply-To`/`References` opcionales, destinatarios exactos (nunca derivados del
  `Reply-To`), `Message-ID` propio devuelto al caller, alternativa en texto plano, imágenes
  inline y adjuntos, y `autoSubmitted` (`Auto-Submitted` + `X-Auto-Response-Suppress: All`)
  para cortar loops con respuestas automáticas.
- **`findByInternetMessageId`, `moveMessage` y `getOrCreateFolder`** (con carpetas ocultas)
  en ambos providers. En Gmail, una "carpeta" es un label.
- **IDs inmutables opcionales en Microsoft:** `immutableIds: true` envía
  `Prefer: IdType="ImmutableId"` en cada llamada a Graph, así el id de un mensaje sobrevive
  a un cambio de carpeta.

### Changed

- **BREAKING — el contrato `EmailProvider` tiene 4 métodos nuevos**
  (`sendConversationMessage`, `findByInternetMessageId`, `moveMessage`,
  `getOrCreateFolder`), y `MessageSummary.cc` y `MessageFull.kind` son campos
  obligatorios nuevos. Quien implemente `EmailProvider` o construya esos tipos (mocks en
  tests) tiene que agregarlos. Quien solo consume la librería no cambia nada. Coordinar con
  synmail y AgentFleet al bumpear el pin.
- **En Microsoft, un reply o forward con `origin` pasa por el flujo de borrador**
  (`createReply`/`createForward` → enviar), aunque no lleve adjuntos: Graph solo acepta el
  header al crear el mensaje, y solo se verificó en las acciones `create*`. Sin `origin` se
  sigue usando el endpoint de un paso.
- **El builder MIME codifica en base64 las partes de texto (HTML y texto plano).** Antes el
  HTML iba sin codificar; las líneas largas y el contenido no ASCII ahora sobreviven a
  cualquier relay.

### Security

- **Gmail: inyección de headers por CR/LF en valores del caller.** Un asunto, destinatario o
  `In-Reply-To` con salto de línea podía agregar headers arbitrarios (p. ej. un `Bcc`) al
  MIME enviado. El builder ahora quita CR/LF de todo valor de header. Preexistente
  (≤ v0.5.1).

## [0.5.1] — 2026-10-13

### Security

- **`MessageFull.internalSender` ahora exige que el propio `From` haya enviado el correo,
  no solo que lo haya enviado alguien de la organización.** En v0.5.0 bastaban los sellos de
  Exchange (`AuthAs: Internal`, `Originating`, `Hosted`), pero esos describen a quien
  *envió* el mensaje, no al dueño del `From`: una regla de buzón que redirige un correo
  externo con `From` falsificado lo reenvía como interno, y un delegado (send on behalf)
  envía con el `From` de otra persona. Ahora además se requiere que `sender` de Graph sea
  igual a `from` y que no haya señales de reenvío o generación automática
  (`X-MS-Exchange-Organization-AutoForwarded`, `X-MS-Exchange-Inbox-Rules-Loop`,
  `X-MS-Exchange-Generated-Message-Source`, `Resent-*`, `Auto-Submitted` distinto de
  `no`); ante cualquier ambigüedad, `false`. No detectable: buzones con permiso *SendAs* o
  apps con `Mail.Send` app-only envían con `sender = from`; esa suplantación la otorga el
  admin del tenant. Impacto: AgentFleet usa este campo para dejar
  que un remitente interno actúe como sí mismo (p. ej. cancelar su tarea programada); con
  v0.5.0 un falso positivo permitía actuar como otra persona. Responde al comentario de
  review en #6. **Follow-up en AgentFleet:** fijar el pin en `v0.5.1`, no `v0.5.0`.

## [0.5.0] — 2026-10-13

### Added

- **`MessageFull.internalSender` (solo Microsoft 365): indica si un correo viene de un
  remitente autenticado de la MISMA organización.** El `From` es falsificable y DKIM/DMARC
  no sirven (Exchange no firma el correo interno: `dkim=none`, `dmarc=none`). Ahora
  `readMessage` devuelve `true` solo si Exchange estampó el mensaje con
  `X-MS-Exchange-Organization-AuthAs: Internal`,
  `X-MS-Exchange-Organization-MessageDirectionality: Originating` y
  `X-MS-Exchange-CrossTenant-FromEntityHeader: Hosted`, cada uno exactamente una vez;
  duplicados, faltantes o valores contradictorios dan `false`. Exchange borra o sobrescribe
  esos headers en el correo entrante, así que no se pueden forjar desde afuera (verificado
  con un envío externo que los traía forjados). En Google queda `undefined`. Retrocompatible
  (campo opcional nuevo). Habilita que AgentFleet actúe sobre pedidos de remitentes internos
  (p. ej. cancelar su propia tarea programada respondiendo al correo del resultado; ver
  AgentFleet#275 y #358). Cierra #6. **Follow-up en AgentFleet:** bumpear el pin a `v0.5.0`.

### Ops

- **Primera suite de tests (`npm test`):** `node:test` vía `tsx`, con fixtures de headers
  anonimizados en `test/fixtures/`; `npm run typecheck` ahora cubre también `test/`.

## [0.4.2] — 2026-10-07

### Fixed

- **Microsoft: los nombres neutros de carpeta (`sent`, `trash`, `spam`, …) ya no
  fallan contra Graph.** `listMessages` y `searchMessages` de
  `MicrosoftEmailProvider` enviaban el `folder` crudo a
  `/mailFolders/{folder}/messages`, y Graph respondía `400 Id is malformed` porque
  solo acepta ids o sus nombres well-known. Ahora se mapean (sin distinguir
  mayúsculas ni espacios alrededor) `sent → sentitems`, `trash → deleteditems`,
  `spam`/`junk → junkemail` y `draft`/`drafts → drafts`; los ids reales y los
  nombres well-known pasan intactos. Google acepta ahora los mismos alias (`draft`,
  `junk`), así que `inbox`, `sent`, `drafts`, `trash` y `spam` funcionan igual en
  ambos providers; `starred`/`important` siguen siendo solo de Gmail. Cierra #5.
  **Follow-up en AgentFleet:** al bumpear el pin, `normalizeGraphFolder`
  (`src/core/email-folder.ts`) puede reducirse a sus alias propios (en español y
  nombres largos como `papelera`, `enviados`, `sent items`); los nombres neutros
  ya los resuelve esta librería.

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
