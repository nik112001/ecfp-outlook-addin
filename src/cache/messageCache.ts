/**
 * messageCache.ts
 *
 * IndexedDB cache layer for per-message carbon footprint results.
 * Keyed by Outlook message ID; backed by the `idb` library (v8).
 *
 * DB name   : "ecfp"
 * DB version: 1
 * Store     : "footprints"  — keyPath: "messageId"
 * Index     : "calculatedAt" — for timestamp-based queries and future cleanup
 */

import { openDB, type IDBPDatabase } from "idb";

// ── Public types ────────────────────────────────────────────────────────────────

/**
 * Shape of a cached footprint record stored in IndexedDB.
 * Mirrors the fields produced by calcSendReceive plus metadata.
 */
export interface CachedFootprint {
  /** Outlook message ID (used as the object-store key). */
  messageId: string;
  /** Total carbon footprint in grams of CO₂ equivalent. */
  gCO2e: number;
  /** Message payload size in megabytes. */
  sizeMB: number;
  /** Total recipient count (To + Cc + Bcc). */
  recipientCount: number;
  /** Whether the message stayed on the corporate intranet. */
  isIntranet: boolean;
  /** Unix timestamp (Date.now()) when the result was calculated. */
  calculatedAt: number;
  /** Formula version string from calcEngine.METHODOLOGY_VERSION. */
  methodologyVersion: string;
}

// ── Internal constants ──────────────────────────────────────────────────────────

const DB_NAME = "ecfp";
const DB_VERSION = 1;
const STORE_NAME = "footprints";
const INDEX_NAME = "calculatedAt";

// ── Singleton database handle ───────────────────────────────────────────────────

/** Module-level promise so we open the database at most once. */
let _dbPromise: Promise<IDBPDatabase> | null = null;

/**
 * Open (or reuse) the IndexedDB database.
 *
 * The first call creates the database and sets up the object store and index.
 * Subsequent calls return the cached promise, ensuring a single connection
 * is shared across all module consumers.
 *
 * @returns A resolved IDBPDatabase instance.
 */
export async function openCache(): Promise<IDBPDatabase> {
  if (_dbPromise === null) {
    _dbPromise = openDB(DB_NAME, DB_VERSION, {
      upgrade(db) {
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          const store = db.createObjectStore(STORE_NAME, {
            keyPath: "messageId",
          });
          // Index on calculatedAt supports cleanup queries and incremental sync.
          store.createIndex(INDEX_NAME, "calculatedAt");
        }
      },
    });
  }
  return _dbPromise;
}

// ── CRUD operations ─────────────────────────────────────────────────────────────

/**
 * Retrieve a cached footprint by Outlook message ID.
 *
 * @param messageId - The Outlook message ID to look up.
 * @returns The matching CachedFootprint, or undefined if not found.
 */
export async function getFootprint(
  messageId: string
): Promise<CachedFootprint | undefined> {
  const db = await openCache();
  return db.get(STORE_NAME, messageId) as Promise<CachedFootprint | undefined>;
}

/**
 * Insert or replace a footprint record.
 *
 * Uses IndexedDB `put` semantics: an existing record with the same messageId
 * is overwritten, making this safe to call on recalculation.
 *
 * @param fp - The footprint record to store.
 */
export async function putFootprint(fp: CachedFootprint): Promise<void> {
  const db = await openCache();
  await db.put(STORE_NAME, fp);
}

/**
 * Delete a footprint record by message ID.
 *
 * No-ops silently if the record does not exist.
 *
 * @param messageId - The Outlook message ID of the record to remove.
 */
export async function deleteFootprint(messageId: string): Promise<void> {
  const db = await openCache();
  await db.delete(STORE_NAME, messageId);
}

/**
 * Remove every footprint record from the store.
 *
 * Intended for user-initiated "clear cache" actions or test teardown.
 */
export async function clearAllFootprints(): Promise<void> {
  const db = await openCache();
  await db.clear(STORE_NAME);
}

/**
 * Return all stored footprint records.
 *
 * Intended for dashboard aggregation; for large mailboxes callers should
 * prefer paginated access via the calculatedAt index in a future iteration.
 *
 * @returns Array of every CachedFootprint in the store (order not guaranteed).
 */
export async function getAllFootprints(): Promise<CachedFootprint[]> {
  const db = await openCache();
  return db.getAll(STORE_NAME) as Promise<CachedFootprint[]>;
}

/**
 * Return footprint records whose calculatedAt timestamp is >= fromTimestamp.
 *
 * Uses the "calculatedAt" index with an IDBKeyRange to avoid a full table scan.
 * Useful for incremental sync: pass the timestamp of the last successful sync
 * to retrieve only records calculated since then.
 *
 * @param fromTimestamp - Inclusive lower bound (Unix ms, e.g. Date.now()).
 * @returns Array of matching CachedFootprint records, ordered by calculatedAt.
 */
export async function getFootprintsSince(
  fromTimestamp: number
): Promise<CachedFootprint[]> {
  const db = await openCache();
  const range = IDBKeyRange.lowerBound(fromTimestamp, false); // inclusive
  return db.getAllFromIndex(
    STORE_NAME,
    INDEX_NAME,
    range
  ) as Promise<CachedFootprint[]>;
}
