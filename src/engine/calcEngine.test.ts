import { describe, it, expect } from "vitest";
import { calcSendReceive, calcStorage, calcEquivalencies, METHODOLOGY_VERSION } from "./calcEngine";

/**
 * Calibration targets from spec §4.5.
 *
 * Formula recap:
 *   gCO2e_msg = (E_base + E_data × S_MB × I_mix) × R_factor
 *   E_base=0.3, E_data=11, I_mix=1.0 (internet) | 0.6 (intranet)
 *   R_factor = min(1 + 0.25×(N−1), 3.0)
 *
 *   gCO2e_storage = S_MB × D_years × E_store  (E_store=10)
 */

describe("calcSendReceive", () => {
  it("short reply: ~0.3 g CO₂e (0.001 MB, 1 recipient, internet)", () => {
    // (0.3 + 11 × 0.001 × 1.0) × 1 = (0.3 + 0.011) × 1 = 0.311
    const result = calcSendReceive({ sizeMB: 0.001, recipientCount: 1, isIntranet: false });
    expect(result.gCO2e).toBeCloseTo(0.311, 2);
    // spec calibration target: ~0.3 g
    expect(result.gCO2e).toBeGreaterThan(0.28);
    expect(result.gCO2e).toBeLessThan(0.5);
  });

  it("image/attachment email: in 11–50 g range (4 MB, 1 recipient, internet)", () => {
    // (0.3 + 11 × 4 × 1.0) × 1 = (0.3 + 44) × 1 = 44.3
    const result = calcSendReceive({ sizeMB: 4, recipientCount: 1, isIntranet: false });
    expect(result.gCO2e).toBeCloseTo(44.3, 1);
    expect(result.gCO2e).toBeGreaterThanOrEqual(11);
    expect(result.gCO2e).toBeLessThanOrEqual(50);
  });

  it("R_factor caps at 3.0 for 10+ recipients", () => {
    // N=10: R_factor = min(1 + 0.25×9, 3.0) = min(3.25, 3.0) = 3.0
    const result = calcSendReceive({ sizeMB: 0.1, recipientCount: 10, isIntranet: false });
    expect(result.breakdown.rFactor).toBe(3.0);

    // N=100: still capped at 3.0
    const resultLarge = calcSendReceive({ sizeMB: 0.1, recipientCount: 100, isIntranet: false });
    expect(resultLarge.breakdown.rFactor).toBe(3.0);
  });

  it("R_factor is correct for small recipient counts below cap", () => {
    // N=1: R_factor = 1 + 0.25×0 = 1.0
    const r1 = calcSendReceive({ sizeMB: 1, recipientCount: 1, isIntranet: false });
    expect(r1.breakdown.rFactor).toBeCloseTo(1.0, 5);

    // N=5: R_factor = 1 + 0.25×4 = 2.0
    const r5 = calcSendReceive({ sizeMB: 1, recipientCount: 5, isIntranet: false });
    expect(r5.breakdown.rFactor).toBeCloseTo(2.0, 5);
  });

  it("intranet discount: same size/recipients gives lower result than internet", () => {
    const internet = calcSendReceive({ sizeMB: 2, recipientCount: 3, isIntranet: false });
    const intranet = calcSendReceive({ sizeMB: 2, recipientCount: 3, isIntranet: true });

    expect(intranet.gCO2e).toBeLessThan(internet.gCO2e);

    // I_mix intranet = 0.6, internet = 1.0
    // data_intranet = 11 × 2 × 0.6 = 13.2
    // data_internet = 11 × 2 × 1.0 = 22.0
    // R_factor(3) = 1 + 0.25×2 = 1.5
    // internet: (0.3 + 22.0) × 1.5 = 33.45
    // intranet: (0.3 + 13.2) × 1.5 = 20.25
    expect(internet.gCO2e).toBeCloseTo(33.45, 2);
    expect(intranet.gCO2e).toBeCloseTo(20.25, 2);
  });

  it("breakdown fields are individually correct", () => {
    const result = calcSendReceive({ sizeMB: 1, recipientCount: 1, isIntranet: false });
    // base = 0.3, data = 11×1×1.0 = 11, rFactor = 1.0
    // total = (0.3 + 11) × 1 = 11.3
    expect(result.breakdown.base).toBeCloseTo(0.3, 5);
    expect(result.breakdown.data).toBeCloseTo(11.0, 5);
    expect(result.breakdown.rFactor).toBeCloseTo(1.0, 5);
    expect(result.gCO2e).toBeCloseTo(11.3, 5);
  });

  it("sizeMB=0 with 1 recipient → result equals E_base (0.3 g)", () => {
    // data term = 11 × 0 × 1.0 = 0; R_factor(1) = 1.0
    // total = (0.3 + 0) × 1.0 = 0.3
    const result = calcSendReceive({ sizeMB: 0, recipientCount: 1, isIntranet: false });
    expect(result.gCO2e).toBeCloseTo(0.3, 5);
    expect(result.breakdown.data).toBeCloseTo(0, 5);
    expect(result.breakdown.base).toBeCloseTo(0.3, 5);
  });

  it("recipientCount=0 produces R_factor=0.75 (sub-1.0, not clamped)", () => {
    // Design choice: recipientCount=0 is allowed through the raw formula.
    // R_factor = min(1 + 0.25×(0−1), 3.0) = min(0.75, 3.0) = 0.75.
    // This yields a result lower than a 1-recipient send, which is semantically
    // correct (message never delivered) and avoids a hidden clamp that could
    // mask misconfigured inputs. Callers should use recipientCount ≥ 1 for real sends.
    const result = calcSendReceive({ sizeMB: 1, recipientCount: 0, isIntranet: false });
    expect(result.breakdown.rFactor).toBeCloseTo(0.75, 5);
    // (0.3 + 11) × 0.75 = 11.3 × 0.75 = 8.475
    expect(result.gCO2e).toBeCloseTo(8.475, 5);
  });

  it("very large attachment: 10 MB, 5 recipients, internet → result in 100–400 g range", () => {
    // R_factor(5) = 1 + 0.25×4 = 2.0
    // data = 11 × 10 × 1.0 = 110
    // total = (0.3 + 110) × 2.0 = 220.6
    const result = calcSendReceive({ sizeMB: 10, recipientCount: 5, isIntranet: false });
    expect(result.gCO2e).toBeCloseTo(220.6, 1);
    expect(result.gCO2e).toBeGreaterThan(100);
    expect(result.gCO2e).toBeLessThan(400);
  });
});

describe("calcStorage", () => {
  it("1 MB × 1 year × E_store(10) = 10 g CO₂e", () => {
    const result = calcStorage({ sizeMB: 1, ageYears: 1 });
    expect(result).toBeCloseTo(10, 5);
  });

  it("scales linearly with size and duration", () => {
    const r1 = calcStorage({ sizeMB: 2, ageYears: 1 });
    const r2 = calcStorage({ sizeMB: 1, ageYears: 2 });
    const r3 = calcStorage({ sizeMB: 2, ageYears: 2 });
    expect(r1).toBeCloseTo(20, 5);
    expect(r2).toBeCloseTo(20, 5);
    expect(r3).toBeCloseTo(40, 5);
  });

  it("returns 0 for 0 MB or 0 years", () => {
    expect(calcStorage({ sizeMB: 0, ageYears: 5 })).toBe(0);
    expect(calcStorage({ sizeMB: 5, ageYears: 0 })).toBe(0);
  });

  it("0.5 MB × 2 years = 10 g CO₂e", () => {
    // 0.5 × 2 × 10 = 10
    const result = calcStorage({ sizeMB: 0.5, ageYears: 2 });
    expect(result).toBeCloseTo(10, 5);
  });
});

describe("calcEquivalencies", () => {
  it("170 g → 1.0 km driven (spec §4.6: 170 g CO₂e per km)", () => {
    const eq = calcEquivalencies(170);
    expect(eq.kmDriven).toBeCloseTo(1.0, 5);
  });

  it("8 g → 1.0 phone charge (spec §4.6: 8 g CO₂e per charge)", () => {
    const eq = calcEquivalencies(8);
    expect(eq.phoneCharges).toBeCloseTo(1.0, 5);
  });

  it("0 g → 0 for both equivalencies", () => {
    const eq = calcEquivalencies(0);
    expect(eq.kmDriven).toBe(0);
    expect(eq.phoneCharges).toBe(0);
  });

  it("scales proportionally for arbitrary values", () => {
    // 340 g → 2 km driven, 42.5 phone charges
    const eq = calcEquivalencies(340);
    expect(eq.kmDriven).toBeCloseTo(2.0, 5);
    expect(eq.phoneCharges).toBeCloseTo(42.5, 5);
  });
});

describe("METHODOLOGY_VERSION", () => {
  it("exports version string '0.3'", () => {
    expect(METHODOLOGY_VERSION).toBe('0.3');
  });
});
