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
  MessageKind,
  AttachmentInfo,
  Folder,
  ListResult,
  Attachment,
  RetrievedAttachment,
  DraftRef,
  // Operation params
  SendParams,
  DraftParams,
  ConversationMessageParams,
  ListParams,
  ListSinceParams,
  ListSinceResult,
  SearchParams,
  ReplyParams,
  ForwardParams,
} from "./types.js";
