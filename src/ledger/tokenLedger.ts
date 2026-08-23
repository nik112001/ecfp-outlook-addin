/**
 * tokenLedger.ts
 *
 * Local IndexedDB ledger for the eCFP monthly token budget system.
 * Tracks debits (email footprint) and credits (cleanup actions) against a
 * configurable monthly budget; computes badge tier from percentage used.
 *
 * DB name   : "ecfp-ledger"   (separate from the "ecfp" footprint cache)
 * DB version: 1
 * Stores:
 *   "config"       — keyPath: "key"           — budget & deduplication state
 *   "transactions" — keyPath: "id" (autoIncr) — ledger history
 */

import { openDB, type IDBPDatabase } from "idb";

// ── Public types ────────────────────────────────────────────────────────────────

/** A single entry in the running ledger. */
export interface LedgerTransaction {
  /** Auto-assigned primary key by IndexedDB. Absent on records not yet stored. */
  id?: number;
  /** Direction of the transaction. */
  type: "debit" | "credit";
  /** Carbon amount in grams CO₂ equivalent. */
  gCO2e: number;
  /** Human-readable reason, e.g. "Email read: msg123" or "Cleanup: deleted 5 messages". */
  reason: string;
  /** Running balance after this transaction (can go negative — real value, not floored). */
  balanceAfter: number;
  /** Unix millisecond timestamp of the transaction. */
  timestamp: number;
  /** ISO year-month slice, e.g. "2026-08". Used for monthly reset queries. */
  month: string;
}

/** Aggregate snapshot of the ledger for the current month. */
export interface LedgerSummary {
  /** Current token balance. Can be negative when over-budget (real, not floored). */
  balance: number;
  /** Configured monthly budget in gCO2e (default: 10 000 g = 10 kg). */
  monthlyBudget: number;
  /** Total gCO2e debited this month. */
  spent: number;
  /** Total gCO2e credited this month (cleanup actions). */
  credited: number;
  /** Percentage of budget consumed: spent / monthlyBudget × 100. */
  percentUsed: number;
  /** Engagement badge tier derived from percentUsed. */
  tier: BadgeTier;
}

/**
 * Engagement badge tier as defined in spec §2 Phase 4.
 *
 * - platinum  : percentUsed <= 50
 * - gold      : percentUsed <= 75
 * - silver    : percentUsed <= 100
 * - over-budget: percentUsed > 100
 */
export type BadgeTier = "platinum" | "gold" | "silver" | "over-budget";

// ── Internal constants ──────────────────────────────────────────────────────────

const DB_NAME = "ecfp-ledger";
const DB_VERSION = 1;
const STORE_CONFIG = "config";
const STORE_TXN = "transactions";

const DEFAULT_MONTHLY_BUDGET = 10_000; // 10 000 g CO2e = 10 kg

/** Config key for the user-configured monthly budget. */
const KEY_BUDGET = "monthlyBudget";
/** Config key for the month string of the last reset (e.g. "2026-08"). */
const KEY_CURRENT_MONTH = "currentMonth";
/** Config key prefix for per-month deduplication sets (appended with YYYY-MM). */
const KEY_DEBITED_IDS_PREFIX = "debitedIds-";

// ── Singleton DB handle ─────────────────────────────────────────────────────────

/** Module-level promise — open the database at most once per process lifetime. */
let _dbPromise: Promise<IDBPDatabase> | null = null;

/**
 * Open (or reuse) the ecfp-ledger IndexedDB database.
 *
 * Creates both object stores on first run. Subsequent calls return the cached
 * promise so all module consumers share a single connection.
 */
export async function openLedger(): Promise<IDBPDatabase> {
  if (_dbPromise === null) {
    _dbPromise = openDB(DB_NAME, DB_VERSION, {
      upgrade(db) {
        if (!db.objectStoreNames.contains(STORE_CONFIG)) {
          db.createObjectStore(STORE_CONFIG, { keyPath: "key" });
        }
        if (!db.objectStoreNames.contains(STORE_TXN)) {
          db.createObjectStore(STORE_TXN, {
            keyPath: "id",
            autoIncrement: true,
          });
        }
      },
    });
  }
  return _dbPromise;
}

// ── Internal helpers ────────────────────────────────────────────────────────────

/** Return the current ISO year-month string, e.g. "2026-08". */
function currentMonth(): string {
  return new Date().toISOString().slice(0, 7);
}

/** Read a config value by key; returns undefined if not present. */
async function getConfig<T>(
  db: IDBPDatabase,
  key: string
): Promise<T | undefined> {
  const row = await db.get(STORE_CONFIG, key);
  return row ? (row.value as T) : undefined;
}

/** Write a config value by key. */
async function setConfig(
  db: IDBPDatabase,
  key: string,
  value: unknown
): Promise<void> {
  await db.put(STORE_CONFIG, { key, value });
}

/** Derive the badge tier from the percentage of budget used. */
function computeTier(percentUsed: number): BadgeTier {
  if (percentUsed <= 50) return "platinum";
  if (percentUsed <= 75) return "gold";
  if (percentUsed <= 100) return "silver";
  return "over-budget";
}

// ── Public API ──────────────────────────────────────────────────────────────────

/**
 * Return the user-configured monthly budget in gCO2e.
 * If no budget has been set, initialises and persists the default (10 000 g).
 */
export async function getMonthlyBudget(): Promise<number> {
  const db = await openLedger();
  const stored = await getConfig<number>(db, KEY_BUDGET);
  if (stored !== undefined) return stored;
  // Initialise the default on first call.
  await setConfig(db, KEY_BUDGET, DEFAULT_MONTHLY_BUDGET);
  return DEFAULT_MONTHLY_BUDGET;
}

/**
 * Persist a new monthly budget value.
 *
 * @param gCO2e - New budget in grams CO₂ equivalent (must be > 0).
 */
export async function setMonthlyBudget(gCO2e: number): Promise<void> {
  const db = await openLedger();
  await setConfig(db, KEY_BUDGET, gCO2e);
}

/**
 * Debit the ledger for a single email's carbon footprint.
 *
 * Idempotent within the current month: if `messageId` has already been debited
 * this month the call is silently ignored, preventing double-counting when the
 * task pane re-opens on the same message.
 *
 * @param messageId - Outlook message ID (used for deduplication).
 * @param gCO2e     - Carbon footprint of the email in grams CO₂ equivalent.
 */
export async function debitEmail(
  messageId: string,
  gCO2e: number
): Promise<void> {
  const db = await openLedger();
  const month = currentMonth();
  const debitedKey = `${KEY_DEBITED_IDS_PREFIX}${month}`;

  // --- Deduplication check ---
  const debitedIds = (await getConfig<string[]>(db, debitedKey)) ?? [];
  if (debitedIds.includes(messageId)) return; // already recorded this month

  // --- Compute new balance ---
  const budget = await getMonthlyBudget();
  // Fetch the last transaction's balanceAfter to get the running balance.
  // If no transactions exist yet, the balance starts at the monthly budget.
  const allTxns = await getMonthTransactionsFromDb(db, month);
  const currentBalance =
    allTxns.length > 0 ? allTxns[allTxns.length - 1].balanceAfter : budget;

  const newBalance = currentBalance - gCO2e;

  // --- Record transaction ---
  const txn: LedgerTransaction = {
    type: "debit",
    gCO2e,
    reason: `Email read: ${messageId}`,
    balanceAfter: newBalance,
    timestamp: Date.now(),
    month,
  };
  await db.add(STORE_TXN, txn);

  // --- Update deduplication set ---
  await setConfig(db, debitedKey, [...debitedIds, messageId]);
}

/**
 * Credit the ledger for a cleanup action (e.g. deleting flagged aged mail).
 *
 * Credits increase the running balance above budget if the user has been
 * cleaning aggressively; this is intentional and reflects the proposal's
 * cleanup incentive mechanic.
 *
 * @param gCO2e       - Carbon saving in grams CO₂ equivalent.
 * @param description - Human-readable description of the cleanup action.
 */
export async function creditCleanup(
  gCO2e: number,
  description: string
): Promise<void> {
  const db = await openLedger();
  const month = currentMonth();

  const budget = await getMonthlyBudget();
  const allTxns = await getMonthTransactionsFromDb(db, month);
  const currentBalance =
    allTxns.length > 0 ? allTxns[allTxns.length - 1].balanceAfter : budget;

  const newBalance = currentBalance + gCO2e;

  const txn: LedgerTransaction = {
    type: "credit",
    gCO2e,
    reason: description,
    balanceAfter: newBalance,
    timestamp: Date.now(),
    month,
  };
  await db.add(STORE_TXN, txn);
}

/**
 * Return an aggregate summary of the ledger for the current month.
 *
 * Automatically triggers a month reset if the stored active-month differs
 * from the current calendar month.
 */
export async function getLedgerSummary(): Promise<LedgerSummary> {
  const db = await openLedger();
  const month = currentMonth();

  // Auto-reset if the calendar month has rolled over.
  const storedMonth = await getConfig<string>(db, KEY_CURRENT_MONTH);
  if (storedMonth !== undefined && storedMonth !== month) {
    await resetForNewMonth();
  }

  const budget = await getMonthlyBudget();
  const allTxns = await getMonthTransactionsFromDb(db, month);

  let spent = 0;
  let credited = 0;
  for (const t of allTxns) {
    if (t.type === "debit") spent += t.gCO2e;
    else credited += t.gCO2e;
  }

  const balance =
    allTxns.length > 0 ? allTxns[allTxns.length - 1].balanceAfter : budget;

  const percentUsed = budget > 0 ? (spent / budget) * 100 : 0;
  const tier = computeTier(percentUsed);

  return { balance, monthlyBudget: budget, spent, credited, percentUsed, tier };
}

/**
 * Return all transactions recorded for the current month, ordered by
 * autoIncrement id (i.e. insertion order).
 */
export async function getMonthTransactions(): Promise<LedgerTransaction[]> {
  const db = await openLedger();
  return getMonthTransactionsFromDb(db, currentMonth());
}

/**
 * Reset the ledger for a new calendar month.
 *
 * Clears the transaction history (the old month's records are removed) and
 * resets the "currentMonth" config key to the current month. The running
 * balance implicitly restarts at the monthly budget because there are no
 * transactions yet.
 *
 * Note: this does NOT clear the debitedIds set from the old month — those keys
 * become stale naturally and a future housekeeping pass can prune them.
 */
export async function resetForNewMonth(): Promise<void> {
  const db = await openLedger();
  await db.clear(STORE_TXN);
  await setConfig(db, KEY_CURRENT_MONTH, currentMonth());
}

// ── Private DB helpers ──────────────────────────────────────────────────────────

/**
 * Retrieve all transactions for a given month string from an already-open DB.
 * Filters in-process because the transactions store has no month index —
 * total volume per month is small so a full scan is acceptable.
 */
async function getMonthTransactionsFromDb(
  db: IDBPDatabase,
  month: string
): Promise<LedgerTransaction[]> {
  const all = (await db.getAll(STORE_TXN)) as LedgerTransaction[];
  return all.filter((t) => t.month === month);
}
