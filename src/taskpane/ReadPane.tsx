import React, { useState, useEffect } from "react";
import {
  Text,
  Card,
  CardHeader,
  Divider,
  Link,
  ProgressBar,
  Spinner,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { calcSendReceive, calcEquivalencies, CalcResult, METHODOLOGY_VERSION } from "../engine/calcEngine";
import { putFootprint } from "../cache/messageCache";

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
  if (gCO2e < 2) return "#107c10"; // green
  if (gCO2e <= 10) return "#c19c00"; // amber
  return "#d13438"; // red
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

    const calculate = () => {
      setLoading(true);
      readOfficeItemData()
        .then((emailData) => {
          if (cancelled) return;
          const result = calcSendReceive(emailData);
          setDisplay({ emailData, result });
          setIsDemo(false);
          setLoading(false);
          const itemId = (Office as any)?.context?.mailbox?.item?.itemId as string | undefined; // eslint-disable-line
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

    // Register ItemChanged so the pane updates when the user clicks a different email
    if (typeof Office !== "undefined" && Office.context?.mailbox) {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => { if (!cancelled) calculate(); }
      );
    }

    calculate();

    return () => {
      cancelled = true;
      if (typeof Office !== "undefined" && Office.context?.mailbox) {
        Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged, () => {});
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
        <div className={styles.spinnerContainer}>
          <Spinner label="Calculating…" size="medium" />
        </div>
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
