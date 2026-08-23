/**
 * mailboxScanner.ts
 *
 * Scans the user's mailbox via Microsoft Graph API, computes per-message
 * carbon footprints, stores results in IndexedDB, and returns aggregated stats.
 *
 * Spec references:
 *   §2 Phase 3     — Cumulative eCFP Dashboard (stored-mail footprint, cleanup nudges)
 *   §3.4 Data Flows — Graph /me/messages paged, incremental sync, 429/Retry-After
 *   §4.4           — Storage formula: gCO2e_storage = S_MB × D_years × E_store
 *   §8.1           — Success criteria (30-day sync < 45 s cold; < 2 s cached)
 *   §9 Risks       — Cap to 2,000 messages / 12 months; incremental extension
 */

import { getGraphClient, GraphMessage } from "./graphClient";
import { getSignedInEmail } from "../auth/msalClient";
import {
  calcSendReceive,
  calcStorage,
  METHODOLOGY_VERSION,
} from "../engine/calcEngine";
import {
  CachedFootprint,
  putFootprint,
  getAllFootprints,
} from "../cache/messageCache";

// ── Public types ────────────────────────────────────────────────────────────────

/**
 * Progress snapshot emitted by scanMailbox during the fetch/compute phases.
 * `total` is null until the first Graph page returns a count estimate.
 */
export interface ScanProgress {
  scanned: number;
  total: number | null; // null until first page returns
  phase: "fetching" | "computing" | "done" | "error";
}

/**
 * A message flagged as a cleanup opportunity (spec §2 Phase 3, §9).
 * Candidates are large (> 5 MB) or old (> 1 year), sorted by storageCO2ePerYear desc.
 */
export interface CleanupCandidate {
  messageId: string;
  subject?: string; // not available under Mail.ReadBasic — omitted from scan
  sizeMB: number;
  ageYears: number;
  storageCO2ePerYear: number; // g CO₂e saved per year if deleted
  receivedDateTime: string;
}

/**
 * Aggregated footprint statistics returned by scanMailbox and aggregateStats.
 */
export interface AggregatedStats {
  /** Month-to-date send/receive footprint in g CO₂e. */
  mtdGCO2e: number;
  /** Year-to-date send/receive footprint in g CO₂e. */
  ytdGCO2e: number;
  /** Ongoing storage footprint for all cached messages in g CO₂e. */
  totalStorageGCO2e: number;
  /** Total number of messages in cache. */
  messageCount: number;
  /** Top 20 cleanup candidates, sorted by storageCO2ePerYear desc (spec §2 Phase 3). */
  cleanupCandidates: CleanupCandidate[];
  /** Unix timestamp (ms) when the scan completed. */
  lastSyncAt: number;
}

// ── Internal constants ──────────────────────────────────────────────────────────

/** Maximum messages to process in a single scan run (spec §9). */
const MAX_MESSAGES = 2000;

/** Graph page size — small enough to keep individual requests fast. */
const PAGE_SIZE = 50;

/** Number of months to look back on first scan (spec §9). */
const LOOKBACK_MONTHS = 12;

/** Emit an onProgress callback every N messages during compute phase. */
const PROGRESS_INTERVAL = 10;

/** Size threshold (MB) above which a message is a cleanup candidate (spec §2 Phase 3). */
const CLEANUP_SIZE_THRESHOLD_MB = 5;

/** Age threshold (years) above which a message is a cleanup candidate (spec §2 Phase 3). */
const CLEANUP_AGE_THRESHOLD_YEARS = 1;

/** Maximum cleanup candidates to include in AggregatedStats. */
const MAX_CLEANUP_CANDIDATES = 20;

// ── Helpers ─────────────────────────────────────────────────────────────────────

/**
 * Extract the domain from an email address (lower-cased).
 * Returns an empty string if the address is malformed.
 */
function emailDomain(address: string): string {
  const at = address.lastIndexOf("@");
  return at >= 0 ? address.slice(at + 1).toLowerCase() : "";
}

/**
 * ISO 8601 timestamp for the start of the day `months` months ago (UTC).
 */
function isoMonthsAgo(months: number): string {
  const d = new Date();
  d.setUTCMonth(d.getUTCMonth() - months);
  d.setUTCHours(0, 0, 0, 0);
  return d.toISOString();
}

/**
 * ISO 8601 timestamp for a Unix millisecond timestamp.
 */
function isoFromMs(ms: number): string {
  return new Date(ms).toISOString();
}

/**
 * Unix timestamp (ms) for midnight UTC on the first day of the current month.
 */
function startOfCurrentMonthMs(): number {
  const d = new Date();
  return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1);
}

/**
 * Unix timestamp (ms) for midnight UTC on 1 Jan of the current year.
 */
function startOfCurrentYearMs(): number {
  return Date.UTC(new Date().getUTCFullYear(), 0, 1);
}

/**
 * Sleep for `seconds` seconds. Used to honour Graph 429 Retry-After.
 */
function sleep(seconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, seconds * 1000));
}

// ── Core scan logic ─────────────────────────────────────────────────────────────

/**
 * Scan the signed-in user's mailbox, compute per-message footprints, persist
 * them to IndexedDB, and return aggregated stats.
 *
 * Design decisions:
 * - Incremental sync: if cached data exists, only fetch messages received after
 *   the most recent cached `calculatedAt`. Falls back to a 12-month window on
 *   first run (spec §9, §3.4).
 * - 429 handling: reads `Retry-After` response header and waits before retrying
 *   the same page (spec §3.4, §9 "Graph throttling" risk).
 * - Cap at MAX_MESSAGES (2,000) per spec §9 to bound scan time.
 * - onProgress fired every PROGRESS_INTERVAL messages and at phase transitions.
 *
 * @param onProgress - Callback for live scan progress updates.
 * @returns Aggregated stats built from the full IndexedDB cache after scan.
 * @throws Error("Not signed in") if no authenticated user is available.
 * @throws Re-throws non-429 Graph errors after setting phase to 'error'.
 */
export async function scanMailbox(
  onProgress: (p: ScanProgress) => void
): Promise<AggregatedStats> {
  // ── 1. Guard: require authenticated user ────────────────────────────────────
  const signedInEmail = getSignedInEmail();
  if (signedInEmail === null) {
    throw new Error("Not signed in");
  }
  const userDomain = emailDomain(signedInEmail);

  // ── 2. Determine incremental sync start timestamp ───────────────────────────
  const existingFootprints = await getAllFootprints();
  const existingIds = new Set(existingFootprints.map((fp) => fp.messageId));

  let filterTimestamp: string;
  if (existingFootprints.length > 0) {
    // Incremental: start from most recent cached calculatedAt (convert ms → ISO)
    const latestCalculatedAt = Math.max(
      ...existingFootprints.map((fp) => fp.calculatedAt)
    );
    filterTimestamp = isoFromMs(latestCalculatedAt);
  } else {
    // First run: look back LOOKBACK_MONTHS (spec §9 cap)
    filterTimestamp = isoMonthsAgo(LOOKBACK_MONTHS);
  }

  // ── 3. Build initial Graph request ──────────────────────────────────────────
  const client = getGraphClient();

  const selectFields =
    "id,size,receivedDateTime,from,toRecipients,ccRecipients,hasAttachments";
  const filter = `receivedDateTime ge ${filterTimestamp}`;
  const orderBy = "receivedDateTime desc";

  // ── 4. Page through results ─────────────────────────────────────────────────
  // `nextUrl` is either null (first page, use SDK builder) or an @odata.nextLink
  // string (subsequent pages, passed directly to client.api()).
  let nextUrl: string | null = null;
  let isFirstPage = true;
  let totalScanned = 0;
  let estimatedTotal: number | null = null;
  const now = Date.now();

  onProgress({ scanned: 0, total: null, phase: "fetching" });

  while (totalScanned < MAX_MESSAGES) {
    // Fetch next page — with 429 retry loop
    let pageData: { value: GraphMessage[]; "@odata.nextLink"?: string };
    let retryCount = 0;
    const MAX_RETRIES = 5;

    while (true) {
      try {
        if (isFirstPage) {
          // Use the fluent SDK builder for the first request
          pageData = await client
            .api("/me/messages")
            .select(selectFields)
            .filter(filter)
            .orderby(orderBy)
            .top(PAGE_SIZE)
            .get();
        } else {
          // Follow @odata.nextLink directly; it already contains all query params
          pageData = await client.api(nextUrl!).get();
        }
        break; // success — exit retry loop
      } catch (err: unknown) {
        const graphError = err as {
          statusCode?: number;
          responseHeaders?: Record<string, string>;
          message?: string;
        };

        if (graphError?.statusCode === 429 && retryCount < MAX_RETRIES) {
          // Honor Retry-After header (spec §3.4, §9)
          const retryAfterHeader =
            graphError.responseHeaders?.["retry-after"] ??
            graphError.responseHeaders?.["Retry-After"] ??
            "5";
          const retryAfterSeconds = parseInt(retryAfterHeader, 10) || 5;
          retryCount++;
          await sleep(retryAfterSeconds);
          continue;
        }

        // Non-429 error or retries exhausted — surface to caller (spec error handling)
        onProgress({ scanned: totalScanned, total: estimatedTotal, phase: "error" });
        throw err;
      }
    }

    const messages: GraphMessage[] = pageData.value ?? [];

    // Graph paged results don't expose a total count by default (that would
    // require $count=true plus ConsistencyLevel:eventual, which is additional
    // overhead). `estimatedTotal` stays null throughout; the dashboard can show
    // an indeterminate progress indicator and update to the final count on done.

    // ── 5. Process each message in the page ──────────────────────────────────
    onProgress({ scanned: totalScanned, total: estimatedTotal, phase: "computing" });

    for (const msg of messages) {
      if (totalScanned >= MAX_MESSAGES) break;

      // Skip messages already in cache (idempotency for incremental sync)
      if (existingIds.has(msg.id)) {
        continue;
      }

      // Convert bytes → MB
      const sizeMB = (msg.size ?? 0) / (1024 * 1024);

      // Count total recipients (To + Cc); Graph doesn't provide Bcc via ReadBasic
      const toCount = msg.toRecipients?.length ?? 0;
      const ccCount = msg.ccRecipients?.length ?? 0;
      const recipientCount = Math.max(toCount + ccCount, 1); // treat 0 as 1 (undeliverable edge case)

      // Intranet detection: all recipients share the sender domain AND sender
      // domain matches the signed-in user's domain (spec §3.2, §4.3 I_mix)
      const senderAddress =
        msg.from?.emailAddress?.address?.toLowerCase() ?? "";
      const senderDomain = emailDomain(senderAddress);

      const allRecipientAddresses = [
        ...(msg.toRecipients?.map((r) => r.emailAddress?.address ?? "") ?? []),
        ...(msg.ccRecipients?.map((r) => r.emailAddress?.address ?? "") ?? []),
      ];

      const isIntranet =
        senderDomain.length > 0 &&
        senderDomain === userDomain &&
        allRecipientAddresses.every(
          (addr) => emailDomain(addr) === senderDomain
        );

      // Compute send/receive footprint
      const calcResult = calcSendReceive({ sizeMB, recipientCount, isIntranet });

      // Build and persist the cache record
      const record: CachedFootprint = {
        messageId: msg.id,
        gCO2e: calcResult.gCO2e,
        sizeMB,
        recipientCount,
        isIntranet,
        calculatedAt: now,
        methodologyVersion: METHODOLOGY_VERSION,
      };

      await putFootprint(record);
      existingIds.add(msg.id);
      totalScanned++;

      // Emit progress every PROGRESS_INTERVAL messages
      if (totalScanned % PROGRESS_INTERVAL === 0) {
        onProgress({ scanned: totalScanned, total: estimatedTotal, phase: "computing" });
      }
    }

    // Mark first page done and advance to next page (or stop)
    isFirstPage = false;
    nextUrl = pageData["@odata.nextLink"] ?? null;
    if (nextUrl === null) break; // no more pages
  }

  // ── 6. Scan complete — aggregate and return ─────────────────────────────────
  onProgress({ scanned: totalScanned, total: totalScanned, phase: "done" });

  return aggregateStats();
}

// ── Aggregation ─────────────────────────────────────────────────────────────────

/**
 * Read all cached footprints from IndexedDB and compute aggregate statistics
 * for the dashboard (spec §2 Phase 3, §5.3).
 *
 * Called automatically at the end of scanMailbox; also exported for the
 * dashboard to refresh stats without re-scanning.
 *
 * Cleanup candidates: messages where ageYears > 1 OR sizeMB > 5 (spec §2 Phase 3).
 * Sorted by storageCO2ePerYear desc; capped at MAX_CLEANUP_CANDIDATES (20).
 *
 * @returns AggregatedStats built from the current IndexedDB contents.
 */
export async function aggregateStats(): Promise<AggregatedStats> {
  const allFootprints = await getAllFootprints();
  const nowMs = Date.now();
  const msPerYear = 365.25 * 24 * 3600 * 1000;

  const mtdStart = startOfCurrentMonthMs();
  const ytdStart = startOfCurrentYearMs();

  let mtdGCO2e = 0;
  let ytdGCO2e = 0;
  let totalStorageGCO2e = 0;
  const candidates: CleanupCandidate[] = [];

  for (const fp of allFootprints) {
    // MTD / YTD send-receive footprint
    if (fp.calculatedAt >= mtdStart) {
      mtdGCO2e += fp.gCO2e;
    }
    if (fp.calculatedAt >= ytdStart) {
      ytdGCO2e += fp.gCO2e;
    }

    // Storage footprint: age in years from when the record was calculated
    // (calculatedAt approximates receivedDateTime for the purposes of this
    // formula — a future iteration can store receivedDateTime directly)
    const ageYears = (nowMs - fp.calculatedAt) / msPerYear;
    const storageGCO2e = calcStorage({ sizeMB: fp.sizeMB, ageYears });
    totalStorageGCO2e += storageGCO2e;

    // Cleanup candidate check (spec §2 Phase 3)
    if (ageYears > CLEANUP_AGE_THRESHOLD_YEARS || fp.sizeMB > CLEANUP_SIZE_THRESHOLD_MB) {
      // storageCO2ePerYear = annual ongoing cost = sizeMB × 1 year × E_store
      // This is the savings per year if the message is deleted.
      const storageCO2ePerYear = calcStorage({ sizeMB: fp.sizeMB, ageYears: 1 });

      candidates.push({
        messageId: fp.messageId,
        // subject is not available under Mail.ReadBasic (spec §3.3 permission scope)
        sizeMB: fp.sizeMB,
        ageYears,
        storageCO2ePerYear,
        // calculatedAt approximates receivedDateTime; stored as ISO string
        receivedDateTime: new Date(fp.calculatedAt).toISOString(),
      });
    }
  }

  // Sort candidates by annual savings desc, cap at MAX_CLEANUP_CANDIDATES
  candidates.sort((a, b) => b.storageCO2ePerYear - a.storageCO2ePerYear);
  const cleanupCandidates = candidates.slice(0, MAX_CLEANUP_CANDIDATES);

  return {
    mtdGCO2e,
    ytdGCO2e,
    totalStorageGCO2e,
    messageCount: allFootprints.length,
    cleanupCandidates,
    lastSyncAt: nowMs,
  };
}
