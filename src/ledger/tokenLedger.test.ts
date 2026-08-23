/**
 * tokenLedger.test.ts
 *
 * Unit tests for the local eCFP token ledger.
 *
 * `fake-indexeddb/auto` patches the global `indexedDB` and related globals
 * before any module under test is imported, giving a fully-functional in-memory
 * IndexedDB environment in Node/Vitest without a browser.
 */

import "fake-indexeddb/auto";

import { beforeEach, describe, expect, it } from "vitest";
import {
  creditCleanup,
  debitEmail,
  getLedgerSummary,
  getMonthlyBudget,
  getMonthTransactions,
  openLedger,
  resetForNewMonth,
  setMonthlyBudget,
  type BadgeTier,
  type LedgerSummary,
  type LedgerTransaction,
} from "./tokenLedger";

// ── Test setup ───────────────────────────────────────────────────────────────────

/**
 * Wipe all stores between tests so each test starts from a clean slate.
 * We clear both stores and re-seed "currentMonth" to avoid auto-reset
 * side-effects from getLedgerSummary detecting a month mismatch.
 */
beforeEach(async () => {
  const db = await openLedger();
  await db.clear("transactions");
  await db.clear("config");
});

// ── openLedger ───────────────────────────────────────────────────────────────────

describe("openLedger", () => {
  it("resolves to a database instance with the expected object stores", async () => {
    const db = await openLedger();
    expect(db).toBeDefined();
    expect(db.objectStoreNames.contains("config")).toBe(true);
    expect(db.objectStoreNames.contains("transactions")).toBe(true);
  });
});

// ── getMonthlyBudget / setMonthlyBudget ──────────────────────────────────────────

describe("getMonthlyBudget", () => {
  it("returns the default budget (10 000 g) when no budget has been set", async () => {
    const budget = await getMonthlyBudget();
    expect(budget).toBe(10_000);
  });

  it("persists the default so a second call returns the same value", async () => {
    await getMonthlyBudget(); // seeds the default
    const budget = await getMonthlyBudget();
    expect(budget).toBe(10_000);
  });
});

describe("setMonthlyBudget / getMonthlyBudget", () => {
  it("round-trips a custom budget value", async () => {
    await setMonthlyBudget(5_000);
    expect(await getMonthlyBudget()).toBe(5_000);
  });

  it("overwrites the budget when called multiple times", async () => {
    await setMonthlyBudget(3_000);
    await setMonthlyBudget(8_500);
    expect(await getMonthlyBudget()).toBe(8_500);
  });

  it("accepts a very large budget", async () => {
    await setMonthlyBudget(1_000_000);
    expect(await getMonthlyBudget()).toBe(1_000_000);
  });
});

// ── debitEmail ───────────────────────────────────────────────────────────────────

describe("debitEmail", () => {
  it("reduces the balance by the debited amount", async () => {
    await debitEmail("msg-001", 500);
    const summary = await getLedgerSummary();
    expect(summary.balance).toBe(10_000 - 500);
    expect(summary.spent).toBe(500);
  });

  it("accumulates multiple debits correctly", async () => {
    await debitEmail("msg-001", 200);
    await debitEmail("msg-002", 300);
    const summary = await getLedgerSummary();
    expect(summary.balance).toBe(10_000 - 500);
    expect(summary.spent).toBe(500);
  });

  it("is idempotent — debiting the same messageId twice does not double-count", async () => {
    await debitEmail("msg-dup", 400);
    await debitEmail("msg-dup", 400); // second call must be ignored
    const summary = await getLedgerSummary();
    expect(summary.balance).toBe(10_000 - 400);
    expect(summary.spent).toBe(400);
  });

  it("records a transaction with correct fields", async () => {
    const before = Date.now();
    await debitEmail("msg-tx", 100);
    const after = Date.now();

    const txns = await getMonthTransactions();
    expect(txns).toHaveLength(1);
    const t = txns[0];
    expect(t.type).toBe("debit");
    expect(t.gCO2e).toBe(100);
    expect(t.reason).toContain("msg-tx");
    expect(t.balanceAfter).toBe(10_000 - 100);
    expect(t.timestamp).toBeGreaterThanOrEqual(before);
    expect(t.timestamp).toBeLessThanOrEqual(after);
    expect(t.month).toMatch(/^\d{4}-\d{2}$/);
  });

  it("allows the balance to go negative when over-budget", async () => {
    await setMonthlyBudget(100);
    await debitEmail("msg-big", 150);
    const summary = await getLedgerSummary();
    expect(summary.balance).toBe(-50);
    expect(summary.tier).toBe("over-budget");
  });
});

// ── creditCleanup ────────────────────────────────────────────────────────────────

describe("creditCleanup", () => {
  it("increases the balance by the credited amount", async () => {
    await debitEmail("msg-001", 1_000); // spend some first
    await creditCleanup(200, "Cleanup: deleted 3 messages");
    const summary = await getLedgerSummary();
    expect(summary.balance).toBe(10_000 - 1_000 + 200);
    expect(summary.credited).toBe(200);
  });

  it("records a credit transaction with correct type and reason", async () => {
    await creditCleanup(500, "Cleanup: deleted 5 messages");
    const txns = await getMonthTransactions();
    expect(txns).toHaveLength(1);
    const t = txns[0];
    expect(t.type).toBe("credit");
    expect(t.gCO2e).toBe(500);
    expect(t.reason).toBe("Cleanup: deleted 5 messages");
  });

  it("can credit without any prior debits (balance exceeds budget)", async () => {
    await creditCleanup(300, "Early cleanup");
    const summary = await getLedgerSummary();
    expect(summary.balance).toBe(10_000 + 300);
    expect(summary.credited).toBe(300);
    expect(summary.spent).toBe(0);
  });

  it("accumulates multiple credits correctly", async () => {
    await creditCleanup(100, "Cleanup A");
    await creditCleanup(200, "Cleanup B");
    const summary = await getLedgerSummary();
    expect(summary.credited).toBe(300);
  });
});

// ── getLedgerSummary ─────────────────────────────────────────────────────────────

describe("getLedgerSummary", () => {
  it("returns default values on an empty ledger", async () => {
    const summary = await getLedgerSummary();
    expect(summary.monthlyBudget).toBe(10_000);
    expect(summary.balance).toBe(10_000);
    expect(summary.spent).toBe(0);
    expect(summary.credited).toBe(0);
    expect(summary.percentUsed).toBe(0);
    expect(summary.tier).toBe("platinum");
  });

  it("computes percentUsed as spent / budget × 100", async () => {
    await setMonthlyBudget(10_000);
    await debitEmail("msg-pct", 2_500);
    const summary = await getLedgerSummary();
    expect(summary.percentUsed).toBeCloseTo(25, 5);
  });

  it("reflects both debits and credits in spent/credited fields independently", async () => {
    await debitEmail("msg-a", 1_000);
    await creditCleanup(400, "Cleanup");
    await debitEmail("msg-b", 600);
    const summary = await getLedgerSummary();
    expect(summary.spent).toBe(1_600);
    expect(summary.credited).toBe(400);
    // balance = 10000 - 1000 + 400 - 600 = 8800
    expect(summary.balance).toBe(8_800);
  });
});

// ── Badge tier logic ─────────────────────────────────────────────────────────────

describe("badge tier computation", () => {
  /**
   * Helper: set a custom budget and debit exactly `fraction` of it,
   * then return the resulting summary.
   */
  async function summaryAtFraction(fraction: number): Promise<LedgerSummary> {
    // Fresh state for each call (beforeEach handles the full reset,
    // but this helper may be called multiple times within one test).
    const db = await openLedger();
    await db.clear("transactions");
    await db.clear("config");

    const budget = 10_000;
    await setMonthlyBudget(budget);
    const amount = budget * fraction;
    if (amount > 0) {
      await debitEmail(`msg-tier-${fraction}`, amount);
    }
    return getLedgerSummary();
  }

  it("platinum  — 50 % used  (percentUsed <= 50)", async () => {
    const s = await summaryAtFraction(0.5); // exactly 50 %
    expect(s.percentUsed).toBeCloseTo(50, 5);
    expect(s.tier).toBe<BadgeTier>("platinum");
  });

  it("gold      — 76 % used  (50 < percentUsed <= 75 → boundary test at 75%)", async () => {
    // 75 % should still be gold (boundary is <= 75 → gold)
    const s75 = await summaryAtFraction(0.75);
    expect(s75.percentUsed).toBeCloseTo(75, 5);
    expect(s75.tier).toBe<BadgeTier>("gold");
  });

  it("silver    — 76 % used  (exceeds gold threshold of 75%)", async () => {
    const s = await summaryAtFraction(0.76);
    expect(s.percentUsed).toBeCloseTo(76, 5);
    expect(s.tier).toBe<BadgeTier>("silver");
  });

  it("silver    — 95 % used  (75 < percentUsed <= 100)", async () => {
    const s = await summaryAtFraction(0.95);
    expect(s.percentUsed).toBeCloseTo(95, 5);
    expect(s.tier).toBe<BadgeTier>("silver");
  });

  it("over-budget — 110 % used  (percentUsed > 100)", async () => {
    const s = await summaryAtFraction(1.1);
    expect(s.percentUsed).toBeCloseTo(110, 5);
    expect(s.tier).toBe<BadgeTier>("over-budget");
  });

  it("platinum  — 0 % used  (empty ledger)", async () => {
    const s = await summaryAtFraction(0);
    expect(s.percentUsed).toBe(0);
    expect(s.tier).toBe<BadgeTier>("platinum");
  });

  it("silver    — exactly 100 % used (boundary: <= 100 → silver, not over-budget)", async () => {
    const s = await summaryAtFraction(1.0);
    expect(s.percentUsed).toBeCloseTo(100, 5);
    expect(s.tier).toBe<BadgeTier>("silver");
  });
});

// ── getMonthTransactions ─────────────────────────────────────────────────────────

describe("getMonthTransactions", () => {
  it("returns an empty array when no transactions have been recorded", async () => {
    const txns = await getMonthTransactions();
    expect(txns).toHaveLength(0);
  });

  it("returns all debits and credits for the current month", async () => {
    await debitEmail("msg-1", 100);
    await debitEmail("msg-2", 200);
    await creditCleanup(50, "Cleanup");
    const txns = await getMonthTransactions();
    expect(txns).toHaveLength(3);
  });

  it("all returned transactions carry the correct month string", async () => {
    await debitEmail("msg-m", 10);
    const txns = await getMonthTransactions();
    const expectedMonth = new Date().toISOString().slice(0, 7);
    expect(txns.every((t) => t.month === expectedMonth)).toBe(true);
  });

  it("transactions are returned in insertion order (by autoIncrement id)", async () => {
    await debitEmail("msg-first", 10);
    await creditCleanup(5, "Cleanup");
    await debitEmail("msg-last", 20);
    const txns = await getMonthTransactions();
    expect(txns[0].type).toBe("debit");
    expect(txns[1].type).toBe("credit");
    expect(txns[2].type).toBe("debit");
  });
});

// ── resetForNewMonth ─────────────────────────────────────────────────────────────

describe("resetForNewMonth", () => {
  it("clears all transactions so the ledger starts fresh", async () => {
    await debitEmail("msg-pre-reset", 500);
    expect((await getMonthTransactions()).length).toBeGreaterThan(0);

    await resetForNewMonth();

    const txns = await getMonthTransactions();
    expect(txns).toHaveLength(0);
  });

  it("after reset the balance equals the monthly budget", async () => {
    await setMonthlyBudget(7_000);
    await debitEmail("msg-pre", 3_000);
    await resetForNewMonth();
    const summary = await getLedgerSummary();
    expect(summary.balance).toBe(7_000);
    expect(summary.spent).toBe(0);
  });

  it("preserves the configured budget across a reset", async () => {
    await setMonthlyBudget(6_000);
    await resetForNewMonth();
    expect(await getMonthlyBudget()).toBe(6_000);
  });
});
