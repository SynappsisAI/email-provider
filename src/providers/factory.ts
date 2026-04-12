import type { EmailProvider, ProviderConfig } from "../types.js";
import { GoogleEmailProvider } from "./google.js";
import { MicrosoftEmailProvider } from "./microsoft.js";

/**
 * Construct an `EmailProvider` from a typed config. The consumer is
 * responsible for sourcing the config (env vars, AWS Secrets Manager,
 * a database row, a CLI flag — whatever fits its model).
 */
export function createEmailProvider(config: ProviderConfig): EmailProvider {
  switch (config.type) {
    case "google":
      return new GoogleEmailProvider(config);
    case "microsoft":
      return new MicrosoftEmailProvider(config);
  }
}
