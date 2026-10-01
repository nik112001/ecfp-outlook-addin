import React, { useState, useEffect } from "react";
import {
  Text,
  Card,
  CardHeader,
  Divider,
  Link,
  ProgressBar,
  Skeleton,
  SkeletonItem,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { calcSendReceive, calcEquivalencies, CalcResult, METHODOLOGY_VERSION } from "../engine/calcEngine";
import { putFootprint } from "../cache/messageCache";
import { debitEmail } from "../ledger/tokenLedger";

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalL,
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
  },
  footprintDisplay: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightBold,
    textAlign: "center",
    padding: `${tokens.spacingVerticalM} 0`,
  },
  factorRow: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  factorLabel: {
    display: "flex",
    justifyContent: "space-between",
  },
  breakdownText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  intranetBadge: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteGreenForeground1,
    fontStyle: "italic",
  },
  equivalency: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  spinnerContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: tokens.spacingVerticalS,
    padding: `${tokens.spacingVerticalL} 0`,
  },
  footer: {
    display: "flex",
    justifyContent: "flex-end",
    marginTop: tokens.spacingVerticalS,
  },
});

// ── Types ─────────────────────────────────────────────────────────────────────

interface EmailData {
  sizeMB: number;
  recipientCount: number;
  isIntranet: boolean;
}

interface DisplayState {
  emailData: EmailData;
  result: CalcResult;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Returns a CSS color string based on gCO2e thresholds from spec §5.1. */
function footprintColor(gCO2e: number): string {
  try {
    const stored = localStorage.getItem("ecfp-thresholds");
    const t = stored ? JSON.parse(stored) as { green: number; amber: number } : { green: 2, amber: 10 };
    if (gCO2e < t.green) return "#2e7d4f";
    if (gCO2e <= t.amber) return "#c19c00";
    return "#d13438";
  } catch {
    if (gCO2e < 2) return "#2e7d4f";
    if (gCO2e <= 10) return "#c19c00";
    return "#d13438";
  }
}

/**
 * Format the driving equivalency from calcEquivalencies into a human-readable string.
 * Displays meters for small values, km for larger ones (spec §4.6).
 */
function formatDrivingEquivalency(gCO2e: number): string {
  const { kmDriven } = calcEquivalencies(gCO2e);
  const meters = Math.round(kmDriven * 1000);
  if (meters < 1000) {
    return `≈ ${meters} m driven`;
  }
  return `≈ ${kmDriven.toFixed(2)} km driven`;
}

/** Extract the domain portion of an email address. */
function extractDomain(email: string): string {
  const parts = email.split("@");
  return parts.length === 2 ? parts[1].toLowerCase() : "";
}

// ── Demo fallback data (used when Office context is unavailable) ───────────────

const DEMO_EMAIL_DATA: EmailData = {
  sizeMB: 0.5,
  recipientCount: 3,
  isIntranet: false,
};

// ── Office.js data-reading logic ──────────────────────────────────────────────

/**
 * Read all metadata from Office.context.mailbox.item.
 * Returns a Promise that resolves to EmailData.
 * Rejects (or resolves with demo data) if Office context is unavailable.
 */
function readOfficeItemData(): Promise<EmailData> {
  return new Promise((resolve, reject) => {
    // eslint-disable-line
    const mailbox =
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.mailbox
        ? Office.context.mailbox
        : null;

    if (!mailbox || !mailbox.item) {
      reject(new Error("Office mailbox context not available"));
      return;
    }

    const item = mailbox.item; // eslint-disable-line

    // Collect recipients from To + Cc
    const toList: Office.EmailAddressDetails[] = item.to ?? [];
    const ccList: Office.EmailAddressDetails[] = item.cc ?? [];
    const allRecipients: Office.EmailAddressDetails[] = [
      ...toList,
      ...ccList,
    ];
    const recipientCount = allRecipients.length;

    // Compute attachment size in MB
    const attachments: Office.AttachmentDetails[] = item.attachments ?? [];
    const attachmentBytes = attachments.reduce(
      (sum, att) => sum + (att.size ?? 0),
      0
    );
    const attachmentMB = attachmentBytes / (1024 * 1024);

    // Determine sender domain and intranet flag
    const senderEmail: string =
      item.from && item.from.emailAddress ? item.from.emailAddress : "";
    const senderDomain = extractDomain(senderEmail);
    const isIntranet =
      senderDomain.length > 0 &&
      recipientCount > 0 &&
      allRecipients.every(
        (r) => extractDomain(r.emailAddress ?? "") === senderDomain
      );

    // Get body as text for a size proxy
    item.body.getAsync(
      Office.CoercionType.Text,
      // eslint-disable-line
      (result: Office.AsyncResult<string>) => {
        let bodyMB = 0;
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          bodyMB = result.value.length / (1024 * 1024);
        }
        const sizeMB = attachmentMB + bodyMB;
        resolve({ sizeMB, recipientCount, isIntranet });
      }
    );
  });
}

// ── Component ─────────────────────────────────────────────────────────────────

export default function ReadPane(): React.ReactElement {
  const styles = useStyles();

  const [loading, setLoading] = useState<boolean>(true);
  const [display, setDisplay] = useState<DisplayState | null>(null);
  const [isDemo, setIsDemo] = useState<boolean>(false);

  useEffect(() => {
    let cancelled = false;
    let lastItemId: string | undefined;

    const calculate = () => {
      if (cancelled) return;
      setLoading(true);
      readOfficeItemData()
        .then((emailData) => {
          if (cancelled) return;
          const result = calcSendReceive(emailData);
          setDisplay({ emailData, result });
          setIsDemo(false);
          setLoading(false);
          const item = Office.context?.mailbox?.item as Office.MessageRead | undefined;
          const itemId = item?.itemId;
          if (itemId) {
            putFootprint({
              messageId: itemId,
              gCO2e: result.gCO2e,
              sizeMB: emailData.sizeMB,
              recipientCount: emailData.recipientCount,
              isIntranet: emailData.isIntranet,
              calculatedAt: Date.now(),
              methodologyVersion: METHODOLOGY_VERSION,
            }).catch(() => {});
            debitEmail(itemId, result.gCO2e).catch(() => {});
          }
        })
        .catch(() => {
          if (cancelled) return;
          const result = calcSendReceive(DEMO_EMAIL_DATA);
          setDisplay({ emailData: DEMO_EMAIL_DATA, result });
          setIsDemo(true);
          setLoading(false);
        });
    };

    // Poll every 800ms — recalculate only when itemId actually changes.
    // ItemChanged event is unreliable in Outlook web for sideloaded add-ins.
    const poll = setInterval(() => {
      if (cancelled) return;
      const currentItem = Office.context?.mailbox?.item as Office.MessageRead | undefined;
      const currentId = currentItem?.itemId;
      if (currentId !== lastItemId) {
        lastItemId = currentId;
        calculate();
      }
    }, 800);

    calculate();

    // Primary: ItemChanged fires when user clicks a different email (requires SupportsPinning)
    if (typeof Office !== 'undefined' && Office.context?.mailbox?.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => {
          lastItemId = undefined; // reset so poll also triggers
          calculate();
        }
      );
    }

    return () => {
      cancelled = true;
      clearInterval(poll);
      if (typeof Office !== 'undefined' && Office.context?.mailbox?.removeHandlerAsync) {
        Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged);
      }
    };
  }, []);

  // ── Loading state ────────────────────────────────────────────────────────────

  if (loading) {
    return (
      <div className={styles.root}>
        <Text size={500} weight="semibold" className={styles.title}>
          Email Carbon Footprint
        </Text>
        <Card>
          <CardHeader header={<Text weight="semibold" size={300}>Estimated footprint</Text>} />
          <Skeleton>
            <SkeletonItem
              size={96}
              style={{ margin: `${tokens.spacingVerticalM} auto`, width: "200px", borderRadius: tokens.borderRadiusMedium }}
            />
          </Skeleton>
        </Card>
        <Divider />
        <Text size={300} weight="semibold">Breakdown</Text>
        <Skeleton>
          <div style={{ display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS }}>
            <SkeletonItem size={16} />
            <SkeletonItem size={8} />
            <SkeletonItem size={8} style={{ width: "75%" }} />
            <SkeletonItem size={8} />
            <SkeletonItem size={8} style={{ width: "55%" }} />
          </div>
        </Skeleton>
      </div>
    );
  }

  // ── Render with data ─────────────────────────────────────────────────────────

  const { emailData, result } = display!;
  const { sizeMB, recipientCount, isIntranet } = emailData;
  const { gCO2e, breakdown } = result;
  const color = footprintColor(gCO2e);

  // Breakdown line: "Base: 0.30 g | Data (X.XX MB): Y.YY g | Recipients (N): ×Z.ZZ"
  const breakdownLine = [
    `Base: ${breakdown.base.toFixed(2)} g`,
    `Data (${sizeMB.toFixed(2)} MB): ${breakdown.data.toFixed(2)} g`,
    `Recipients (${recipientCount}): ×${breakdown.rFactor.toFixed(2)}`,
  ].join(" | ");

  const drivingLine = formatDrivingEquivalency(gCO2e);

  // Progress bar values: normalise each factor contribution to 0–1 for display.
  // Use gCO2e as denominator; each bar shows what fraction that term contributes.
  const totalForBar = gCO2e > 0 ? gCO2e : 1;
  const barFactors = [
    { label: "Base", value: (breakdown.base * breakdown.rFactor) / totalForBar },
    { label: "Data", value: (breakdown.data * breakdown.rFactor) / totalForBar },
  ];

  return (
    <div className={styles.root}>
      <Text size={500} weight="semibold" className={styles.title}>
        Email Carbon Footprint
        {isDemo && (
          <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginLeft: "8px" }}>
            (demo)
          </Text>
        )}
      </Text>

      <Card>
        <CardHeader
          header={
            <Text weight="semibold" size={300}>
              Estimated footprint
            </Text>
          }
        />
        <Text
          className={styles.footprintDisplay}
          style={{ color }}
        >
          🌱 {gCO2e.toFixed(2)} g CO₂e
        </Text>
        {isIntranet && (
          <Text className={styles.intranetBadge}>
            Intranet discount applied
          </Text>
        )}
      </Card>

      <Divider />

      <Text size={300} weight="semibold">
        Breakdown
      </Text>

      <Text className={styles.breakdownText}>{breakdownLine}</Text>

      {barFactors.map((factor) => (
        <div key={factor.label} className={styles.factorRow}>
          <div className={styles.factorLabel}>
            <Text size={200}>{factor.label}</Text>
            <Text size={200}>
              {factor.label === "Base"
                ? `${(breakdown.base * breakdown.rFactor).toFixed(2)} g`
                : `${(breakdown.data * breakdown.rFactor).toFixed(2)} g`}
            </Text>
          </div>
          <ProgressBar
            value={factor.value}
            color={color === "#107c10" ? "success" : color === "#c19c00" ? "warning" : "error"}
          />
        </div>
      ))}

      <Text className={styles.equivalency}>{drivingLine}</Text>

      <Divider />

      <div className={styles.footer}>
        <Link href="#" onClick={(e) => e.preventDefault()}>
          How is this calculated?
        </Link>
      </div>
    </div>
  );
}
