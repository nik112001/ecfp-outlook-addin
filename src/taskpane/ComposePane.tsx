import React, { useState, useEffect, useRef, useCallback } from "react";
import {
  Text,
  Card,
  CardHeader,
  Divider,
  makeStyles,
  tokens,
  Spinner,
} from "@fluentui/react-components";
import {
  calcSendReceive,
  calcEquivalencies,
} from "../engine/calcEngine";

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalL,
  },
  titleRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  footprintDisplay: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightBold,
    textAlign: "center",
    padding: `${tokens.spacingVerticalM} 0`,
  },
  metricRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: `${tokens.spacingVerticalXS} 0`,
  },
  intranetLine: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteGreenForeground1,
    fontStyle: "italic",
  },
  equivalency: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  remediationBox: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalS,
    backgroundColor: tokens.colorStatusDangerBackground1,
    borderRadius: tokens.borderRadiusMedium,
  },
  remediationItem: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorStatusDangerForeground1,
  },
  disclaimer: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground4,
    fontStyle: "italic",
    textAlign: "center",
    marginTop: tokens.spacingVerticalS,
  },
});

// ── Types ─────────────────────────────────────────────────────────────────────

interface CFFState {
  gCO2e: number;
  sizeMB: number;
  recipientCount: number;
  attachmentCount: number;
  attachmentMB: number;
  bodyMB: number;
  isIntranet: boolean;
  color: "green" | "amber" | "red";
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Returns the color tier based on configurable gCO2e thresholds (spec §5.1). */
function colorTier(gCO2e: number): "green" | "amber" | "red" {
  try {
    const stored = localStorage.getItem("ecfp-thresholds");
    const t = stored ? JSON.parse(stored) as { green: number; amber: number } : { green: 2, amber: 10 };
    if (gCO2e < t.green) return "green";
    if (gCO2e <= t.amber) return "amber";
    return "red";
  } catch {
    if (gCO2e < 2) return "green";
    if (gCO2e <= 10) return "amber";
    return "red";
  }
}

/** Maps a color tier to a CSS color string. */
function colorValue(tier: "green" | "amber" | "red"): string {
  if (tier === "green") return "#107c10";
  if (tier === "amber") return "#c19c00";
  return "#d13438";
}

/** Extract the domain from an email address, lower-cased. */
function extractDomain(email: string): string {
  const at = email.indexOf("@");
  return at !== -1 ? email.slice(at + 1).toLowerCase() : "";
}

/**
 * Format the driving equivalency into a human-readable string.
 * Shows meters for small values, km for larger ones (spec §4.6).
 */
function formatDrivingEquivalency(gCO2e: number): string {
  const { kmDriven } = calcEquivalencies(gCO2e);
  const meters = Math.round(kmDriven * 1000);
  if (meters < 1000) {
    return `≈ ${meters} m driven`;
  }
  return `≈ ${kmDriven.toFixed(2)} km driven`;
}

// ── Demo fallback ─────────────────────────────────────────────────────────────

const DEMO_STATE: CFFState = (() => {
  const sizeMB = 0.3;
  const recipientCount = 2;
  const isIntranet = false;
  const { gCO2e } = calcSendReceive({ sizeMB, recipientCount, isIntranet });
  return {
    gCO2e,
    sizeMB,
    recipientCount,
    attachmentCount: 0,
    attachmentMB: 0,
    bodyMB: sizeMB,
    isIntranet,
    color: colorTier(gCO2e),
  };
})();

// ── Office context probe ──────────────────────────────────────────────────────

/** Returns true if we have a valid compose-mode Office mailbox context. */
function hasOfficeContext(): boolean {
  return (
    typeof Office !== "undefined" &&
    Office.context != null &&
    Office.context.mailbox != null &&
    Office.context.mailbox.item != null
  );
}

// ── Component ─────────────────────────────────────────────────────────────────

export default function ComposePane(): React.ReactElement {
  const styles = useStyles();

  const [cff, setCff] = useState<CFFState>(DEMO_STATE);
  const [isDemo, setIsDemo] = useState<boolean>(true);
  const [loading, setLoading] = useState<boolean>(true);

  // Keep a ref to the polling interval so we can clear it on unmount.
  const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null);

  // In-flight guard to prevent overlapping recalculate calls.
  const calculatingRef = useRef<boolean>(false);

  // ── recalculate ─────────────────────────────────────────────────────────────

  const recalculate = useCallback((): void => {
    if (!hasOfficeContext()) return;
    if (calculatingRef.current) return;
    calculatingRef.current = true;
    setLoading(true);

    const item = Office.context.mailbox.item!; // eslint-disable-line

    // Wrap getAsync callbacks in Promises for parallel execution.
    const toPromise = new Promise<Office.EmailAddressDetails[]>((resolve) => {
      item.to.getAsync((result: Office.AsyncResult<Office.EmailAddressDetails[]>) => { // eslint-disable-line
        resolve(
          result.status === Office.AsyncResultStatus.Succeeded
            ? result.value
            : []
        );
      });
    });

    const ccPromise = new Promise<Office.EmailAddressDetails[]>((resolve) => {
      item.cc.getAsync((result: Office.AsyncResult<Office.EmailAddressDetails[]>) => { // eslint-disable-line
        resolve(
          result.status === Office.AsyncResultStatus.Succeeded
            ? result.value
            : []
        );
      });
    });

    const bodyPromise = new Promise<number>((resolve) => {
      item.body.getAsync(
        Office.CoercionType.Text,
        (result: Office.AsyncResult<string>) => { // eslint-disable-line
          const bodyMB =
            result.status === Office.AsyncResultStatus.Succeeded
              ? result.value.length / (1024 * 1024)
              : 0;
          resolve(bodyMB);
        }
      );
    });

    Promise.all([toPromise, ccPromise, bodyPromise])
      .then(
        ([toList, ccList, bodyMB]) => {
          // Attachment size is synchronous in compose mode.
          const attachments: Office.AttachmentDetails[] =
            item.attachments ?? [];
          const attachmentCount = attachments.length;
          const attachmentMB =
            attachments.reduce((sum, att) => sum + (att.size ?? 0), 0) /
            (1024 * 1024);

          const allRecipients = [...toList, ...ccList];
          const recipientCount = allRecipients.length;

          // Determine intranet: all recipients share the composer's domain.
          const composerEmail: string =
            Office.context.mailbox.userProfile?.emailAddress ?? "";
          const composerDomain = extractDomain(composerEmail);
          const isIntranet =
            composerDomain.length > 0 &&
            recipientCount > 0 &&
            allRecipients.every(
              (r) =>
                extractDomain(r.emailAddress ?? "") === composerDomain
            );

          const sizeMB = attachmentMB + bodyMB;
          const { gCO2e } = calcSendReceive({
            sizeMB,
            recipientCount,
            isIntranet,
          });

          setCff({
            gCO2e,
            sizeMB,
            recipientCount,
            attachmentCount,
            attachmentMB,
            bodyMB,
            isIntranet,
            color: colorTier(gCO2e),
          });
        }
      )
      .finally(() => {
        calculatingRef.current = false;
        setLoading(false);
      });
  }, []);

  // ── Mount / unmount ─────────────────────────────────────────────────────────

  useEffect(() => {
    if (!hasOfficeContext()) {
      // No Office context — stay on demo state.
      setIsDemo(true);
      return;
    }

    setIsDemo(false);

    const item = Office.context.mailbox.item!; // eslint-disable-line

    // Initial calculation.
    recalculate();

    // Event listeners for recipient and attachment changes.
    item.addHandlerAsync(
      Office.EventType.RecipientsChanged,
      (_ev: Office.RecipientsChangedEventArgs) => { // eslint-disable-line
        recalculate();
      }
    );
    item.addHandlerAsync(
      Office.EventType.AttachmentsChanged,
      (_ev: Office.AttachmentsChangedEventArgs) => { // eslint-disable-line
        recalculate();
      }
    );

    // Body polling every 2 s (spec §3.4).
    intervalRef.current = setInterval(() => {
      recalculate();
    }, 2000);

    return () => {
      // Cleanup: remove event handlers and stop polling.
      if (intervalRef.current !== null) {
        clearInterval(intervalRef.current);
        intervalRef.current = null;
      }
      item.removeHandlerAsync(Office.EventType.RecipientsChanged);
      item.removeHandlerAsync(Office.EventType.AttachmentsChanged);
    };
  }, [recalculate]);

  // ── Derived display values ──────────────────────────────────────────────────

  const {
    gCO2e,
    recipientCount,
    attachmentCount,
    attachmentMB,
    bodyMB,
    isIntranet,
    color,
  } = cff;

  const cssColor = colorValue(color);
  const drivingLine = formatDrivingEquivalency(gCO2e);

  const showAttachmentRemediation = color === "red" && attachmentCount > 0;
  const showCcRemediation = color === "red" && recipientCount > 5;

  // ── Render ──────────────────────────────────────────────────────────────────

  return (
    <div className={styles.root}>
      {/* Title */}
      <div className={styles.titleRow}>
        <Text size={500} weight="semibold">
          Current Email Footprint (CFF)
        </Text>
        {isDemo && (
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            (demo)
          </Text>
        )}
      </div>

      {/* Large color-coded footprint number */}
      <Card>
        <CardHeader
          header={
            <Text weight="semibold" size={300}>
              Live estimate
            </Text>
          }
        />
        {loading && !isDemo ? (
          <Spinner label="Calculating…" size="medium" />
        ) : (
          <Text
            className={styles.footprintDisplay}
            style={{ color: cssColor }}
          >
            🌿 {gCO2e.toFixed(2)} g CO₂e
          </Text>
        )}
      </Card>

      <Divider />

      {/* Factor rows */}
      <Text size={300} weight="semibold">
        Factors
      </Text>

      <div className={styles.metricRow}>
        <Text size={200}>Recipients</Text>
        <Text size={200} weight="semibold">
          {recipientCount}
        </Text>
      </div>

      <div className={styles.metricRow}>
        <Text size={200}>Attachments</Text>
        <Text size={200} weight="semibold">
          {attachmentCount} ({attachmentMB.toFixed(2)} MB)
        </Text>
      </div>

      <div className={styles.metricRow}>
        <Text size={200}>Body</Text>
        <Text size={200} weight="semibold">
          ~{bodyMB.toFixed(2)} MB
        </Text>
      </div>

      {/* Intranet discount notice */}
      {isIntranet && (
        <Text className={styles.intranetLine}>
          Intranet discount applied (×0.6)
        </Text>
      )}

      {/* Equivalency line */}
      <Text className={styles.equivalency}>{drivingLine}</Text>

      <Divider />

      {/* Remediation suggestions — only shown when red */}
      {(showAttachmentRemediation || showCcRemediation) && (
        <div className={styles.remediationBox}>
          {showAttachmentRemediation && (
            <Text className={styles.remediationItem}>
              💡 Consider sharing via OneDrive instead of attaching files.
            </Text>
          )}
          {showCcRemediation && (
            <Text className={styles.remediationItem}>
              💡 Consider trimming your CC list.
            </Text>
          )}
        </div>
      )}

      {/* Disclaimer */}
      <Text className={styles.disclaimer}>
        Live estimate · never blocks send
      </Text>
    </div>
  );
}
