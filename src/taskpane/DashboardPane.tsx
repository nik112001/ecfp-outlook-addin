import React, { useState, useEffect } from "react";
import {
  Text,
  Card,
  CardHeader,
  Button,
  Spinner,
  Divider,
  Link,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from "recharts";
import {
  AggregatedStats,
  CleanupCandidate,
  scanMailbox,
  aggregateStats,
  ScanProgress,
} from "../graph/mailboxScanner";
import { calcEquivalencies, METHODOLOGY_VERSION } from "../engine/calcEngine";

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalL,
  },
  header: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  headerActions: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalS,
  },
  spinnerRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  statRow: {
    display: "flex",
    gap: tokens.spacingHorizontalM,
    flexWrap: "wrap",
  },
  statBlock: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    flex: "1 1 0",
    minWidth: "120px",
  },
  statValue: {
    fontWeight: tokens.fontWeightBold,
    fontSize: tokens.fontSizeBase500,
  },
  statLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  equivalency: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: tokens.spacingVerticalXS,
  },
  chartContainer: {
    marginTop: tokens.spacingVerticalS,
  },
  cleanupRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: `${tokens.spacingVerticalXS} 0`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  cleanupMeta: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  cleanupMetaText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  noCleanup: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
  cardFooter: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground4,
    fontStyle: "italic",
    marginTop: tokens.spacingVerticalS,
  },
  mutedCard: {
    backgroundColor: tokens.colorNeutralBackground3,
  },
  footnote: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
    flexWrap: "wrap",
    marginTop: tokens.spacingVerticalS,
  },
  footnoteText: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground4,
  },
  errorText: {
    color: tokens.colorStatusDangerForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  lastSynced: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
});

// ── Types ─────────────────────────────────────────────────────────────────────

type DashboardStatus = "idle" | "scanning" | "ready" | "error";

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Format a gram value for display.
 * Below 1000 g → "X.X g CO₂e"; 1000 g and above → "X.XX kg CO₂e".
 */
function formatCO2e(g: number): string {
  if (g < 1000) {
    return `${g.toFixed(1)} g CO₂e`;
  }
  return `${(g / 1000).toFixed(2)} kg CO₂e`;
}

/**
 * Returns a human-readable relative time string for a Unix timestamp (ms).
 * e.g. "just now", "3 minutes ago", "2 hours ago", "5 days ago".
 */
function relativeTime(ts: number): string {
  const diffMs = Date.now() - ts;
  const diffSeconds = Math.floor(diffMs / 1000);
  if (diffSeconds < 60) return "just now";
  const diffMinutes = Math.floor(diffSeconds / 60);
  if (diffMinutes < 60) return `${diffMinutes} minute${diffMinutes === 1 ? "" : "s"} ago`;
  const diffHours = Math.floor(diffMinutes / 60);
  if (diffHours < 24) return `${diffHours} hour${diffHours === 1 ? "" : "s"} ago`;
  const diffDays = Math.floor(diffHours / 24);
  return `${diffDays} day${diffDays === 1 ? "" : "s"} ago`;
}

/**
 * Returns the abbreviated month name for a zero-based month index.
 * e.g. 0 → "Jan", 11 → "Dec".
 */
function monthName(monthIndex: number): string {
  const names = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
  ];
  return names[monthIndex] ?? "?";
}

/** Returns a CSS color for the monthly threshold (spec §5.3). */
function monthlyColor(mtdG: number): string {
  if (mtdG < 5000) return "#107c10"; // green
  if (mtdG <= 15000) return "#c19c00"; // amber
  return "#d13438"; // red
}

// ── Component ─────────────────────────────────────────────────────────────────

export default function DashboardPane(): React.ReactElement {
  const styles = useStyles();

  const [status, setStatus] = useState<DashboardStatus>("idle");
  const [stats, setStats] = useState<AggregatedStats | null>(null);
  const [progress, setProgress] = useState<ScanProgress | null>(null);
  const [error, setError] = useState<string | null>(null);
  // Track when data was last loaded so we can display a relative timestamp.
  const [lastSyncedAt, setLastSyncedAt] = useState<number | null>(null);

  // On mount: attempt to load cached aggregated stats from IndexedDB.
  // aggregateStats always resolves (returns zero-value stats when cache is empty),
  // so we check messageCount > 0 to distinguish "has data" from "no data yet".
  useEffect(() => {
    aggregateStats()
      .then((cached: AggregatedStats) => {
        if (cached.messageCount > 0) {
          setStats(cached);
          setLastSyncedAt(Date.now());
          setStatus("ready");
        }
        // messageCount === 0 → no cached data yet, remain 'idle' for scan CTA.
      })
      .catch(() => {
        // Non-fatal: cache read failure just leaves us in idle state.
      });
  }, []);

  // ── Scan handler ─────────────────────────────────────────────────────────────

  async function handleScan(): Promise<void> {
    setStatus("scanning");
    setError(null);
    try {
      const result = await scanMailbox((p: ScanProgress) => setProgress(p));
      setStats(result);
      setLastSyncedAt(Date.now());
      setStatus("ready");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Scan failed");
      setStatus("error");
    }
  }

  // ── Derived values ────────────────────────────────────────────────────────────

  const currentMonthIndex = new Date().getMonth(); // 0-based
  const currentMonthNumber = currentMonthIndex + 1; // 1-based (1–12)

  const mtdG = stats?.mtdGCO2e ?? 0;
  const ytdG = stats?.ytdGCO2e ?? 0;

  // Monthly average estimate: ytdG / months elapsed so far this year.
  const monthlyAvgG = currentMonthNumber > 0 ? ytdG / currentMonthNumber : ytdG;

  const chartData = [
    { name: `${monthName(currentMonthIndex)} (MTD)`, value: Math.round(mtdG) },
    { name: "Yr avg/mo", value: Math.round(monthlyAvgG) },
  ];

  const { kmDriven } = calcEquivalencies(ytdG);
  const equivalencyLine = `≈ ${kmDriven.toFixed(1)} km driven this year`;

  const mtdColor = monthlyColor(mtdG);

  // ── Render ────────────────────────────────────────────────────────────────────

  return (
    <div className={styles.root}>

      {/* ── 1. Header ─────────────────────────────────────────────────────────── */}
      <div className={styles.header}>
        <Text size={600} weight="semibold">
          eCFP Dashboard
        </Text>
        <Text size={300} style={{ color: tokens.colorNeutralForeground2 }}>
          Your email carbon footprint
        </Text>

        <div className={styles.headerActions}>
          {status === "idle" && (
            <Button
              appearance="primary"
              onClick={() => { void handleScan(); }}
            >
              Scan my mailbox
            </Button>
          )}

          {status === "scanning" && (
            <div className={styles.spinnerRow}>
              <Spinner size="tiny" />
              <Text size={200}>
                Scanning&hellip;{" "}
                {progress !== null
                  ? `${progress.scanned} messages processed`
                  : ""}
              </Text>
            </div>
          )}

          {status === "ready" && (
            <>
              <Button
                appearance="subtle"
                size="small"
                onClick={() => { void handleScan(); }}
              >
                Refresh
              </Button>
              {lastSyncedAt !== null && (
                <Text className={styles.lastSynced}>
                  Last synced: {relativeTime(lastSyncedAt)}
                </Text>
              )}
            </>
          )}

          {status === "error" && error !== null && (
            <Text className={styles.errorText}>{error}</Text>
          )}
        </div>
      </div>

      <Divider />

      {/* ── 2. Cumulative eCFP card ───────────────────────────────────────────── */}
      {status === "ready" && stats !== null && (
        <Card>
          <CardHeader
            header={
              <Text weight="semibold" size={300}>
                Cumulative eCFP
              </Text>
            }
          />

          <div className={styles.statRow}>
            {/* Month-to-date */}
            <div className={styles.statBlock}>
              <Text
                className={styles.statValue}
                style={{ color: mtdColor }}
              >
                {formatCO2e(mtdG)}
              </Text>
              <Text className={styles.statLabel}>
                This month
              </Text>
            </div>

            {/* Year-to-date */}
            <div className={styles.statBlock}>
              <Text className={styles.statValue}>
                {formatCO2e(ytdG)}
              </Text>
              <Text className={styles.statLabel}>
                This year (YTD)
              </Text>
            </div>
          </div>

          <Text className={styles.equivalency}>
            {equivalencyLine}
          </Text>
        </Card>
      )}

      {/* ── 3. Trend chart ────────────────────────────────────────────────────── */}
      {status === "ready" && stats !== null && (
        <Card>
          <CardHeader
            header={
              <Text weight="semibold" size={300}>
                Trend
              </Text>
            }
          />
          <div className={styles.chartContainer}>
            <ResponsiveContainer width="100%" height={160}>
              <BarChart
                data={chartData}
                margin={{ top: 4, right: 8, left: 0, bottom: 4 }}
              >
                <CartesianGrid strokeDasharray="3 3" vertical={false} />
                <XAxis
                  dataKey="name"
                  tick={{ fontSize: 11 }}
                  tickLine={false}
                  axisLine={false}
                />
                <YAxis
                  label={{
                    value: "g CO₂e",
                    angle: -90,
                    position: "insideLeft",
                    style: { fontSize: 10 },
                    offset: 8,
                  }}
                  tick={{ fontSize: 10 }}
                  tickLine={false}
                  axisLine={false}
                  width={56}
                />
                <Tooltip
                  formatter={(value: number) => [`${value} g CO₂e`, "Footprint"]}
                />
                <Bar dataKey="value" fill={tokens.colorBrandBackground} radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>
      )}

      {/* ── 4. Storage footprint card ─────────────────────────────────────────── */}
      {status === "ready" && stats !== null && (
        <Card>
          <CardHeader
            header={
              <Text weight="semibold" size={300}>
                Stored mail footprint
              </Text>
            }
          />

          <Text size={400} weight="semibold">
            {formatCO2e(stats.totalStorageGCO2e)} ongoing storage cost
          </Text>
          <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
            Based on {stats.messageCount} messages in cache
          </Text>

          <Divider style={{ margin: `${tokens.spacingVerticalS} 0` }} />

          {stats.cleanupCandidates.length > 0 ? (
            <>
              {stats.cleanupCandidates.slice(0, 5).map((candidate: CleanupCandidate, idx: number) => (
                <div key={idx} className={styles.cleanupRow}>
                  <div className={styles.cleanupMeta}>
                    <Text size={200} weight="semibold">
                      {candidate.sizeMB.toFixed(1)} MB
                    </Text>
                    <Text className={styles.cleanupMetaText}>
                      {candidate.ageYears.toFixed(1)} yr old &nbsp;&middot;&nbsp; saves ~{formatCO2e(candidate.storageCO2ePerYear)}/yr
                    </Text>
                  </div>
                  <Button
                    size="small"
                    appearance="subtle"
                    onClick={() => {
                      // TODO: actual deletion requires Mail.ReadWrite scope — post-POC
                    }}
                  >
                    Delete
                  </Button>
                </div>
              ))}
            </>
          ) : (
            <Text className={styles.noCleanup}>
              No cleanup candidates found.
            </Text>
          )}

          <Text className={styles.cardFooter}>
            Deleting flagged mail saves ongoing storage energy (spec §4.4)
          </Text>
        </Card>
      )}

      {/* ── 5. Token balance stub ─────────────────────────────────────────────── */}
      {status === "ready" && (
        <Card className={styles.mutedCard}>
          <CardHeader
            header={
              <Text weight="semibold" size={300}>
                eCFP Token Balance
              </Text>
            }
          />
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            Token ledger coming in M4. Budget tracking and badge tiers will
            appear here.
          </Text>
        </Card>
      )}

      {/* ── 6. Methodology footnote ───────────────────────────────────────────── */}
      <div className={styles.footnote}>
        <Text className={styles.footnoteText}>
          Methodology v{METHODOLOGY_VERSION} &nbsp;&middot;&nbsp; Endpoint-side POC &nbsp;&middot;&nbsp; No data leaves your device
        </Text>
        <Link
          href="#"
          onClick={(e: React.MouseEvent<HTMLAnchorElement>) => e.preventDefault()}
          style={{ fontSize: tokens.fontSizeBase100 }}
        >
          Learn more
        </Link>
      </div>

    </div>
  );
}
