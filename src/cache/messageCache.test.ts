/**
 * messageCache.test.ts
 *
 * Unit tests for the IndexedDB cache layer.
 *
 * `fake-indexeddb/auto` patches the global `indexedDB` and related globals
 * before any module under test is imported, giving a fully functional in-memory
 * IndexedDB environment in Node/Vitest without a browser.
 */

import "fake-indexeddb/auto";

import { beforeEach, describe, expect, it } from "vitest";
import { METHODOLOGY_VERSION } from "../engine/calcEngine";
import {
  clearAllFootprints,
  deleteFootprint,
  getAllFootprints,
  getFootprint,
  getFootprintsSince,
  openCache,
  putFootprint,
  type CachedFootprint,
} from "./messageCache";

// ── Helpers ─────────────────────────────────────────────────────────────────────

/** Build a minimal valid CachedFootprint for the given messageId. */
function makeFp(
  messageId: string,
  overrides: Partial<CachedFootprint> = {}
): CachedFootprint {
  return {
    messageId,
    gCO2e: 4.2,
    sizeMB: 0.3,
    recipientCount: 2,
    isIntranet: false,
    calculatedAt: Date.now(),
    methodologyVersion: METHODOLOGY_VERSION,
    ...overrides,
  };
}

// ── Test setup ───────────────────────────────────────────────────────────────────

beforeEach(async () => {
  // Start each test with a clean store so tests remain independent.
  await clearAllFootprints();
});

// ── Tests ────────────────────────────────────────────────────────────────────────

describe("openCache", () => {
  it("resolves to a database instance", async () => {
    const db = await openCache();
    expect(db).toBeDefined();
    expect(db.objectStoreNames.contains("footprints")).toBe(true);
  });
});

describe("putFootprint / getFootprint", () => {
  it("stores a record and retrieves it by messageId", async () => {
    const fp = makeFp("msg-001");
    await putFootprint(fp);

    const retrieved = await getFootprint("msg-001");
    expect(retrieved).toBeDefined();
    expect(retrieved!.messageId).toBe("msg-001");
    expect(retrieved!.gCO2e).toBe(fp.gCO2e);
    expect(retrieved!.sizeMB).toBe(fp.sizeMB);
    expect(retrieved!.recipientCount).toBe(fp.recipientCount);
    expect(retrieved!.isIntranet).toBe(fp.isIntranet);
    expect(retrieved!.calculatedAt).toBe(fp.calculatedAt);
    expect(retrieved!.methodologyVersion).toBe(fp.methodologyVersion);
  });

  it("overwrites an existing record when put is called again with the same messageId", async () => {
    await putFootprint(makeFp("msg-002", { gCO2e: 1.0 }));
    await putFootprint(makeFp("msg-002", { gCO2e: 9.9 })); // overwrite

    const retrieved = await getFootprint("msg-002");
    expect(retrieved!.gCO2e).toBe(9.9);
  });
});

describe("getFootprint — nonexistent ID", () => {
  it("returns undefined for an ID that was never stored", async () => {
    const result = await getFootprint("does-not-exist");
    expect(result).toBeUndefined();
  });
});

describe("deleteFootprint", () => {
  it("removes the record so subsequent get returns undefined", async () => {
    await putFootprint(makeFp("msg-003"));
    await deleteFootprint("msg-003");

    const result = await getFootprint("msg-003");
    expect(result).toBeUndefined();
  });

  it("does not throw when deleting a nonexistent ID", async () => {
    await expect(deleteFootprint("never-stored")).resolves.toBeUndefined();
  });
});

describe("clearAllFootprints / getAllFootprints", () => {
  it("getAllFootprints returns every stored record", async () => {
    await putFootprint(makeFp("msg-a"));
    await putFootprint(makeFp("msg-b"));
    await putFootprint(makeFp("msg-c"));

    const all = await getAllFootprints();
    expect(all).toHaveLength(3);
    const ids = all.map((f) => f.messageId).sort();
    expect(ids).toEqual(["msg-a", "msg-b", "msg-c"]);
  });

  it("getAllFootprints returns an empty array after clearAll", async () => {
    await putFootprint(makeFp("msg-x"));
    await putFootprint(makeFp("msg-y"));

    await clearAllFootprints();

    const all = await getAllFootprints();
    expect(all).toHaveLength(0);
  });

  it("clearAll on an already-empty store does not throw", async () => {
    await expect(clearAllFootprints()).resolves.toBeUndefined();
    const all = await getAllFootprints();
    expect(all).toHaveLength(0);
  });
});

describe("getFootprintsSince", () => {
  it("returns only records with calculatedAt >= the given timestamp", async () => {
    const old = 1_000_000; // epoch ms — safely in the past
    const recent = 2_000_000;
    const boundary = 1_500_000; // midpoint used as query lower bound

    await putFootprint(makeFp("msg-old", { calculatedAt: old }));
    await putFootprint(makeFp("msg-recent", { calculatedAt: recent }));

    const results = await getFootprintsSince(boundary);

    expect(results).toHaveLength(1);
    expect(results[0].messageId).toBe("msg-recent");
  });

  it("includes a record whose calculatedAt exactly equals fromTimestamp (inclusive bound)", async () => {
    const ts = 5_000_000;
    await putFootprint(makeFp("msg-exact", { calculatedAt: ts }));
    await putFootprint(makeFp("msg-before", { calculatedAt: ts - 1 }));

    const results = await getFootprintsSince(ts);

    expect(results).toHaveLength(1);
    expect(results[0].messageId).toBe("msg-exact");
  });

  it("returns an empty array when no records match the timestamp filter", async () => {
    await putFootprint(makeFp("msg-stale", { calculatedAt: 100 }));

    const results = await getFootprintsSince(9_999_999);
    expect(results).toHaveLength(0);
  });

  it("returns all records when fromTimestamp is 0", async () => {
    await putFootprint(makeFp("msg-p", { calculatedAt: 1 }));
    await putFootprint(makeFp("msg-q", { calculatedAt: 2 }));

    const results = await getFootprintsSince(0);
    expect(results).toHaveLength(2);
  });
});
