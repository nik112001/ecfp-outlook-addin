/**
 * Microsoft Graph client factory.
 *
 * Wraps `@microsoft/microsoft-graph-client` with an auth provider that
 * delegates token acquisition to `msalClient.getAccessToken()`.  Every Graph
 * request will transparently refresh or re-acquire the token as needed (silent
 * → popup fallback is handled inside `getAccessToken`).
 *
 * Spec refs: §3.2 (Read-mode data — "authoritative size on Graph response"),
 *            §3.3 (Permissions & Scopes — Mail.ReadBasic only).
 */

import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "../auth/msalClient";

// ---------------------------------------------------------------------------
// Graph client factory
// ---------------------------------------------------------------------------

/**
 * Creates an authenticated Microsoft Graph `Client` instance.
 *
 * The `authProvider` callback is invoked by the SDK before every API call, so
 * the token is always fresh (MSAL handles caching and renewal internally).
 *
 * Usage:
 * ```typescript
 * const client = getGraphClient();
 * const messages = await client
 *   .api("/me/messages")
 *   .select("id,size,receivedDateTime,from,toRecipients,ccRecipients,hasAttachments")
 *   .get();
 * ```
 */
export function getGraphClient(): Client {
  return Client.init({
    authProvider: async (done) => {
      try {
        const accessToken = await getAccessToken();
        done(null, accessToken);
      } catch (err) {
        const error =
          err instanceof Error ? err : new Error(String(err));
        done(error, null);
      }
    },
  });
}

// ---------------------------------------------------------------------------
// Typed message interface
// ---------------------------------------------------------------------------

/**
 * Represents the subset of Graph message fields selected during a mailbox scan.
 *
 * Only fields available via `Mail.ReadBasic` are included.  Message bodies are
 * deliberately excluded (spec §3.3: "no bodies").
 *
 * Corresponds to the `$select` projection used in the dashboard scan:
 * `id,size,receivedDateTime,from,toRecipients,ccRecipients,hasAttachments`
 */
export interface GraphMessage {
  /** Graph message ID (opaque string). */
  id: string;

  /** Total message size in bytes, including attachments. */
  size: number;

  /** ISO 8601 timestamp of when the message was received. */
  receivedDateTime: string;

  /** Sender address (from the `From` header). */
  from: {
    emailAddress: {
      address: string;
    };
  };

  /** Primary recipients (To: line). */
  toRecipients: Array<{
    emailAddress: {
      address: string;
    };
  }>;

  /** Carbon-copy recipients (Cc: line). */
  ccRecipients: Array<{
    emailAddress: {
      address: string;
    };
  }>;

  /**
   * `true` if the message has one or more file attachments.
   * Used as a quick filter before computing the attachment-size term in the
   * carbon formula (spec §4.3, S_MB factor).
   */
  hasAttachments: boolean;
}
