// Public surface for @synappsis/email-provider.

export { createEmailProvider } from "./providers/factory.js";
export { GoogleEmailProvider } from "./providers/google.js";
export { MicrosoftEmailProvider } from "./providers/microsoft.js";

export type {
  // Provider contract
  EmailProvider,
  // Configuration
  ProviderConfig,
  GoogleProviderConfig,
  MicrosoftProviderConfig,
  // Domain types
  EmailAddress,
  MessageSummary,
  MessageFull,
  AttachmentInfo,
  Folder,
  ListResult,
  Attachment,
  RetrievedAttachment,
  // Operation params
  SendParams,
  ListParams,
  SearchParams,
  ReplyParams,
  ForwardParams,
} from "./types.js";
