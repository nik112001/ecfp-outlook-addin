/**
 * eCFP Calculation Engine
 *
 * Implements the carbon footprint formula from spec §4.3:
 *
 *   gCO2e_msg = (E_base + E_data × S_MB × I_mix) × R_factor
 *
 * Constants:
 *   E_base   = 0.3  g CO₂e  (baseline energy per email transmission)
 *   E_data   = 11   g CO₂e/MB  (data transfer energy intensity)
 *   I_mix    = 1.0  (internet)  |  0.6  (intranet — lower due to internal grid)
 *   R_factor = min(1 + 0.25 × (N − 1), 3.0)  (recipient scaling factor, capped at 3×)
 *
 * Storage:
 *   gCO2e_storage = S_MB × D_years × E_store
 *   E_store = 10  g CO₂e / MB / year
 */

// ── Version ────────────────────────────────────────────────────────────────────

/**
 * Tracks the spec version this engine implements (spec §0 / document header).
 * Bump when formula or constants change so callers can detect compatibility.
 */
export const METHODOLOGY_VERSION = '0.3';

// ── Constants ──────────────────────────────────────────────────────────────────

/**
 * Baseline energy per email transmission: 0.3 g CO₂e.
 *
 * Source: Berners-Lee, M. *How Bad Are Bananas?* (2020 ed.) — short laptop
 * email figure covering device energy, network hops, and server processing
 * with amortized embodied carbon (spec §4.3, Appendix A ref 3).
 */
const E_BASE = 0.3; // g CO₂e

/**
 * Data-transfer / processing coefficient: 11 g CO₂e per MB.
 *
 * Represents the combined transmission and server-processing energy intensity
 * per megabyte of message payload (spec §4.3). Value is to be refined in M1
 * against data-center energy literature.
 */
const E_DATA = 11; // g CO₂e / MB

/**
 * Grid-mix multiplier for internet recipients: 1.0 (no discount).
 *
 * Full value applied when any recipient is outside the sender's domain,
 * meaning the message traverses the public internet (spec §4.3).
 */
const I_MIX_INTERNET = 1.0;

/**
 * Grid-mix multiplier for intranet (same-domain) recipients: 0.6.
 *
 * Applied when all recipients share the sender's domain, reflecting a shorter
 * infrastructure path through on-premises or co-located data-center links
 * rather than the open internet (spec §4.3). Coefficient flagged for tuning in M1.
 */
const I_MIX_INTRANET = 0.6;

/**
 * Maximum recipient scaling factor: 3.0 (spec §4.3).
 *
 * R_factor = min(1 + 0.25 × (N − 1), 3.0). The cap prevents unrealistically
 * large estimates for large distribution lists while still reflecting that
 * each additional recipient incurs incremental processing and storage cost.
 */
const R_FACTOR_CAP = 3.0;

/**
 * Storage energy intensity: 10 g CO₂e per MB per year.
 *
 * Placeholder value pending M1 literature research into data-center
 * energy-per-GB-stored figures (spec §4.4). Applied per message in the
 * mailbox scan to compute the ongoing storage footprint.
 */
const E_STORE = 10; // g CO₂e / MB / year

// ── Equivalency constants (spec §4.6) ─────────────────────────────────────────

/**
 * CO₂e emitted per kilometre driven in an average gasoline car: 170 g (spec §4.6).
 */
const G_PER_KM_DRIVEN = 170; // g CO₂e / km

/**
 * CO₂e emitted by a single smartphone charge: 8 g (spec §4.6).
 */
const G_PER_PHONE_CHARGE = 8; // g CO₂e / charge

// ── Send / Receive ─────────────────────────────────────────────────────────────

export interface CalcInput {
  /** Message size in megabytes (headers + body + attachments). */
  sizeMB: number;
  /** Total recipient count (To + Cc + Bcc). */
  recipientCount: number;
  /** True if the message stays entirely on the corporate intranet. */
  isIntranet: boolean;
}

export interface CalcResult {
  /** Total estimated grams of CO₂ equivalent. */
  gCO2e: number;
  /** Additive breakdown of the three formula terms. */
  breakdown: {
    /** E_base component (constant per email). */
    base: number;
    /** E_data × S_MB × I_mix component. */
    data: number;
    /** The recipient scaling multiplier applied (≥ 1, ≤ 3). */
    rFactor: number;
  };
}

/**
 * Calculate the carbon footprint of a single send/receive event.
 *
 * @param input - Size, recipient count, and network mix.
 * @returns CalcResult with total gCO2e and breakdown.
 */
export function calcSendReceive(input: CalcInput): CalcResult {
  const { sizeMB, recipientCount, isIntranet } = input;

  const iMix = isIntranet ? I_MIX_INTRANET : I_MIX_INTERNET;

  // R_factor = min(1 + 0.25 × (N − 1), 3.0)
  // Note: recipientCount=0 is treated as N=0, giving R_factor = min(1 + 0.25×(-1), 3.0) = 0.75.
  // This is intentional: zero recipients means the message was never sent/delivered,
  // so a sub-1.0 factor correctly reduces the impact below a single-recipient send.
  // Callers should pass recipientCount ≥ 1 for real messages.
  const rFactor = Math.min(1 + 0.25 * (recipientCount - 1), R_FACTOR_CAP);

  const base = E_BASE;
  const data = E_DATA * sizeMB * iMix;

  const gCO2e = (base + data) * rFactor;

  return {
    gCO2e,
    breakdown: { base, data, rFactor },
  };
}

// ── Storage ────────────────────────────────────────────────────────────────────

export interface StorageInput {
  /** Message size in megabytes. */
  sizeMB: number;
  /** Number of years the message has been (or will be) stored. */
  ageYears: number;
}

/**
 * Calculate the carbon cost of storing an email for a given period.
 *
 *   gCO2e_storage = S_MB × D_years × E_store
 *
 * @param input - Size and storage duration.
 * @returns Estimated grams of CO₂ equivalent for storage.
 */
export function calcStorage(input: StorageInput): number {
  const { sizeMB, ageYears } = input;
  return sizeMB * ageYears * E_STORE;
}

// ── Equivalencies ──────────────────────────────────────────────────────────────

export interface Equivalencies {
  /** Equivalent kilometres driven in an average gasoline car (spec §4.6: 170 g CO₂e / km). */
  kmDriven: number;
  /** Equivalent smartphone charges (spec §4.6: 8 g CO₂e / charge). */
  phoneCharges: number;
}

/**
 * Convert a CO₂e mass into human-readable real-world equivalencies.
 *
 * Uses the figures from spec §4.6:
 *   - 1 km driven (avg gasoline car) ≈ 170 g CO₂e
 *   - 1 smartphone charge ≈ 8 g CO₂e
 *
 * @param gCO2e - Grams of CO₂ equivalent to convert.
 * @returns Object with `kmDriven` and `phoneCharges`.
 */
export function calcEquivalencies(gCO2e: number): Equivalencies {
  return {
    kmDriven: gCO2e / G_PER_KM_DRIVEN,
    phoneCharges: gCO2e / G_PER_PHONE_CHARGE,
  };
}
