// Semantic carbon footprint colors — matches Microsoft Fluent UI palette
export const CARBON_COLORS = {
  green: "#2e7d4f",   // forest green — warmer than Fluent default, approved design direction
  amber: "#c19c00",   // tokens.colorStatusWarningForeground1 equivalent
  red: "#d13438",     // tokens.colorStatusDangerForeground1 equivalent
} as const;

export type CarbonColorKey = keyof typeof CARBON_COLORS;

// Returns the appropriate color hex for a given gCO2e value and thresholds
export function getCarbonColor(
  gCO2e: number,
  thresholds: { green: number; amber: number } = { green: 2, amber: 10 }
): string {
  if (gCO2e < thresholds.green) return CARBON_COLORS.green;
  if (gCO2e <= thresholds.amber) return CARBON_COLORS.amber;
  return CARBON_COLORS.red;
}

// Reads thresholds from localStorage, falls back to defaults
export function getStoredThresholds(): { green: number; amber: number } {
  try {
    const stored = localStorage.getItem("ecfp-thresholds");
    if (stored) return JSON.parse(stored) as { green: number; amber: number };
  } catch { /* ignore */ }
  return { green: 2, amber: 10 };
}
