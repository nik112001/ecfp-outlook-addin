import React, { useState, useEffect } from "react";
import {
  Text,
  Card,
  CardHeader,
  Divider,
  Link,
  Button,
  Input,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { getLedgerSummary, setMonthlyBudget, LedgerSummary, BadgeTier } from "../ledger/tokenLedger";
import { signOut } from "../auth/msalClient";

// ── Threshold helpers (exported for use by other panes) ───────────────────────

export interface Thresholds {
  green: number;
  amber: number;
}

const THRESHOLDS_KEY = "ecfp-thresholds";
const DEFAULT_THRESHOLDS: Thresholds = { green: 2, amber: 10 };

export function getThresholds(): Thresholds {
  try {
    const raw = localStorage.getItem(THRESHOLDS_KEY);
    if (!raw) return DEFAULT_THRESHOLDS;
    const parsed = JSON.parse(raw) as Partial<Thresholds>;
    return {
      green: typeof parsed.green === "number" ? parsed.green : DEFAULT_THRESHOLDS.green,
      amber: typeof parsed.amber === "number" ? parsed.amber : DEFAULT_THRESHOLDS.amber,
    };
  } catch {
    return DEFAULT_THRESHOLDS;
  }
}

export function saveThresholds(t: Thresholds): void {
  localStorage.setItem(THRESHOLDS_KEY, JSON.stringify(t));
}

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalL,
  },
  pageTitle: {
    fontWeight: tokens.fontWeightSemibold,
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
  },
  fieldRow: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  fieldLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  inputRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  inputUnit: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    whiteSpace: "nowrap",
  },
  staticText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  savedConfirmation: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteGreenForeground1,
  },
  budgetKgLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  badgeTier: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  badgeUsage: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  privacyNote: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
  footer: {
    display: "flex",
    justifyContent: "flex-start",
    marginTop: tokens.spacingVerticalS,
  },
  signedOutText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
});

// ── Badge helpers ─────────────────────────────────────────────────────────────

function badgeEmoji(tier: BadgeTier): string {
  switch (tier) {
    case "platinum":    return "🏆 Platinum — using ≤50% of your monthly budget";
    case "gold":        return "🥇 Gold — using ≤75% of your monthly budget";
    case "silver":      return "🥈 Silver — using ≤100% of your monthly budget";
    case "over-budget": return "🔴 Over budget — consider cleanup actions";
    default:            return "—";
  }
}

// ── Component ─────────────────────────────────────────────────────────────────

export default function SettingsPane(): React.ReactElement {
  const styles = useStyles();

  // ── Threshold state ──────────────────────────────────────────────────────────
  const initialThresholds = getThresholds();
  const [greenVal, setGreenVal] = useState<number>(initialThresholds.green);
  const [amberVal, setAmberVal] = useState<number>(initialThresholds.amber);
  const [thresholdSaved, setThresholdSaved] = useState<boolean>(false);

  // ── Budget state ─────────────────────────────────────────────────────────────
  const [budgetVal, setBudgetVal] = useState<number>(10000);
  const [budgetSaved, setBudgetSaved] = useState<boolean>(false);

  // ── Ledger/badge state ───────────────────────────────────────────────────────
  const [ledgerSummary, setLedgerSummary] = useState<LedgerSummary | null>(null);

  // ── Clear-data state ─────────────────────────────────────────────────────────
  const [dataCleared, setDataCleared] = useState<boolean>(false);

  // ── Sign-out state ───────────────────────────────────────────────────────────
  const [signedOut, setSignedOut] = useState<boolean>(false);

  // ── On mount: load ledger summary ────────────────────────────────────────────
  useEffect(() => {
    let cancelled = false;
    getLedgerSummary()
      .then((summary) => {
        if (!cancelled) {
          setLedgerSummary(summary);
          setBudgetVal(summary.monthlyBudget);
        }
      })
      .catch(() => {
        // Non-fatal: leave ledgerSummary as null
      });
    return () => {
      cancelled = true;
    };
  }, []);

  // ── Handlers ─────────────────────────────────────────────────────────────────

  function handleSaveThresholds(): void {
    saveThresholds({ green: greenVal, amber: amberVal });
    setThresholdSaved(true);
    setTimeout(() => setThresholdSaved(false), 2000);
  }

  function handleSaveBudget(): void {
    setMonthlyBudget(budgetVal);
    setBudgetSaved(true);
    setTimeout(() => setBudgetSaved(false), 2000);
  }

  async function handleClearData(): Promise<void> {
    // Only clear eCFP-specific data, not MSAL auth tokens
    localStorage.removeItem("ecfp-thresholds");
    // Clear IndexedDB stores
    await indexedDB.deleteDatabase("ecfp");
    await indexedDB.deleteDatabase("ecfp-ledger");
    // Reset threshold UI state to defaults
    setGreenVal(DEFAULT_THRESHOLDS.green);
    setAmberVal(DEFAULT_THRESHOLDS.amber);
    // Show brief confirmation
    setDataCleared(true);
    setTimeout(() => setDataCleared(false), 3000);
  }

  async function handleSignOut(): Promise<void> {
    try {
      await signOut();
      setSignedOut(true);
    } catch {
      // Best-effort: if signOut throws, still show signed-out state
      setSignedOut(true);
    }
  }

  // ── Render ───────────────────────────────────────────────────────────────────

  const budgetKg = (budgetVal / 1000).toFixed(1);
  const usagePct =
    ledgerSummary && ledgerSummary.monthlyBudget > 0
      ? Math.round((ledgerSummary.spent / ledgerSummary.monthlyBudget) * 100)
      : null;

  return (
    <div className={styles.root}>
      <Text size={500} weight="semibold" className={styles.pageTitle}>
        Settings
      </Text>

      {/* ── 1. Footprint thresholds ─────────────────────────────────────────── */}
      <Card>
        <CardHeader
          header={
            <Text weight="semibold" size={300} className={styles.sectionTitle}>
              Footprint thresholds
            </Text>
          }
        />
        <div className={styles.section}>
          {/* Green threshold */}
          <div className={styles.fieldRow}>
            <Text className={styles.fieldLabel}>Green up to</Text>
            <div className={styles.inputRow}>
              <Input
                type="number"
                value={String(greenVal)}
                min={0}
                onChange={(_e, data) => {
                  const v = parseFloat(data.value);
                  if (!isNaN(v) && v >= 0) setGreenVal(v);
                }}
                style={{ width: "80px" }}
              />
              <Text className={styles.inputUnit}>g CO₂e</Text>
            </div>
          </div>

          {/* Amber threshold */}
          <div className={styles.fieldRow}>
            <Text className={styles.fieldLabel}>Amber up to</Text>
            <div className={styles.inputRow}>
              <Input
                type="number"
                value={String(amberVal)}
                min={0}
                onChange={(_e, data) => {
                  const v = parseFloat(data.value);
                  if (!isNaN(v) && v >= 0) setAmberVal(v);
                }}
                style={{ width: "80px" }}
              />
              <Text className={styles.inputUnit}>g CO₂e</Text>
            </div>
          </div>

          {/* Red — static */}
          <Text className={styles.staticText}>
            Red: above {amberVal} g CO₂e
          </Text>

          <div className={styles.inputRow}>
            <Button appearance="primary" onClick={handleSaveThresholds}>
              Save thresholds
            </Button>
            {thresholdSaved && (
              <Text className={styles.savedConfirmation}>Saved ✓</Text>
            )}
          </div>
        </div>
      </Card>

      <Divider />

      {/* ── 2. Monthly token budget ─────────────────────────────────────────── */}
      <Card>
        <CardHeader
          header={
            <Text weight="semibold" size={300} className={styles.sectionTitle}>
              Monthly token budget
            </Text>
          }
        />
        <div className={styles.section}>
          <div className={styles.fieldRow}>
            <Text className={styles.fieldLabel}>Monthly budget</Text>
            <div className={styles.inputRow}>
              <Input
                type="number"
                value={String(budgetVal)}
                min={0}
                onChange={(_e, data) => {
                  const v = parseFloat(data.value);
                  if (!isNaN(v) && v >= 0) setBudgetVal(v);
                }}
                style={{ width: "100px" }}
              />
              <Text className={styles.inputUnit}>g CO₂e</Text>
            </div>
            <Text className={styles.budgetKgLabel}>= {budgetKg} kg CO₂e / month</Text>
          </div>

          <div className={styles.inputRow}>
            <Button appearance="primary" onClick={handleSaveBudget}>
              Save budget
            </Button>
            {budgetSaved && (
              <Text className={styles.savedConfirmation}>Saved ✓</Text>
            )}
          </div>
        </div>
      </Card>

      <Divider />

      {/* ── 3. Badge tier ───────────────────────────────────────────────────── */}
      <Card>
        <CardHeader
          header={
            <Text weight="semibold" size={300} className={styles.sectionTitle}>
              Your badge tier
            </Text>
          }
        />
        <div className={styles.section}>
          {ledgerSummary ? (
            <>
              <Text className={styles.badgeTier}>
                {badgeEmoji(ledgerSummary.tier)}
              </Text>
              {usagePct !== null && (
                <Text className={styles.badgeUsage}>
                  {usagePct}% of monthly budget used
                </Text>
              )}
            </>
          ) : (
            <Text className={styles.staticText}>Loading badge data…</Text>
          )}
        </div>
      </Card>

      <Divider />

      {/* ── 4. Data & privacy ───────────────────────────────────────────────── */}
      <Card>
        <CardHeader
          header={
            <Text weight="semibold" size={300} className={styles.sectionTitle}>
              Data &amp; privacy
            </Text>
          }
        />
        <div className={styles.section}>
          <Text className={styles.privacyNote}>
            All data is stored locally on this device. Nothing is sent to any server.
          </Text>

          <Dialog>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary">Clear all local data</Button>
            </DialogTrigger>
            <DialogSurface>
              <DialogBody>
                <DialogTitle>Clear all local data?</DialogTitle>
                <DialogContent>
                  This will delete all cached footprints and ledger history. Your Microsoft sign-in will be preserved. Continue?
                </DialogContent>
                <DialogActions>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Cancel</Button>
                  </DialogTrigger>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="primary" onClick={() => void handleClearData()}>
                      Clear data
                    </Button>
                  </DialogTrigger>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>
          {dataCleared && (
            <Text className={styles.savedConfirmation}>Data cleared.</Text>
          )}
        </div>
      </Card>

      <Divider />

      {/* ── 5. Microsoft account ────────────────────────────────────────────── */}
      <Card>
        <CardHeader
          header={
            <Text weight="semibold" size={300} className={styles.sectionTitle}>
              Microsoft account
            </Text>
          }
        />
        <div className={styles.section}>
          {signedOut ? (
            <Text className={styles.signedOutText}>Signed out.</Text>
          ) : (
            <Button appearance="secondary" onClick={() => void handleSignOut()}>
              Sign out
            </Button>
          )}
        </div>
      </Card>

      <Divider />

      {/* ── 6. Footer link ──────────────────────────────────────────────────── */}
      <div className={styles.footer}>
        <Link
          onClick={() => {
            window.location.search = "?mode=methodology";
          }}
        >
          View methodology
        </Link>
      </div>
    </div>
  );
}
