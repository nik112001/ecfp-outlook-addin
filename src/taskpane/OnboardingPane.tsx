import React, { useState } from "react";
import {
  Text,
  Card,
  Button,
  Divider,
  makeStyles,
  tokens,
} from "@fluentui/react-components";

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
    padding: tokens.spacingHorizontalL,
    minHeight: "100%",
  },
  stepIndicator: {
    color: tokens.colorNeutralForeground3,
  },
  stepContent: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    flex: 1,
  },
  iconDisplay: {
    fontSize: "3rem",
    textAlign: "center" as const,
    padding: `${tokens.spacingVerticalM} 0`,
  },
  bodyText: {
    color: tokens.colorNeutralForeground2,
  },
  featureCardsContainer: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  featureCard: {
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingVerticalS,
  },
  featureEmoji: {
    fontSize: tokens.fontSizeHero700,
    minWidth: "2rem",
    textAlign: "center" as const,
  },
  featureDescription: {
    color: tokens.colorNeutralForeground2,
  },
  buttonRow: {
    display: "flex",
    flexDirection: "row",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap" as const,
  },
  buttonRowSpaced: {
    display: "flex",
    flexDirection: "row",
    justifyContent: "space-between",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap" as const,
  },
  buttonRowStep3: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
});

// ── Feature card data ─────────────────────────────────────────────────────────

interface FeatureCardData {
  emoji: string;
  description: string;
}

const FEATURE_CARDS: FeatureCardData[] = [
  { emoji: "📧", description: "Read pane — footprint of every email you open" },
  { emoji: "✏️", description: "Compose pane — live footprint as you write" },
  { emoji: "📊", description: "Dashboard — your cumulative monthly footprint" },
];

// ── Step sub-components ───────────────────────────────────────────────────────

interface StepProps {
  onNext: () => void;
  onBack?: () => void;
}

function Step1Privacy({ onNext }: StepProps): React.ReactElement {
  const styles = useStyles();

  return (
    <div className={styles.stepContent}>
      <div className={styles.iconDisplay}>🌱</div>

      <Text size={600} weight="bold">
        Welcome to eCFP
      </Text>

      <Text size={300} className={styles.bodyText}>
        eCFP measures the carbon footprint of your emails — right inside Outlook.
      </Text>
      <Text size={300} className={styles.bodyText}>
        Everything runs locally on your device. No email content ever leaves your machine.
      </Text>
      <Text size={300} className={styles.bodyText}>
        We only read metadata: recipient count, message size, and attachment info.
      </Text>

      <div className={styles.buttonRow}>
        <Button appearance="primary" onClick={onNext}>
          Next →
        </Button>
      </div>
    </div>
  );
}

function Step2Features({ onNext, onBack }: StepProps): React.ReactElement {
  const styles = useStyles();

  return (
    <div className={styles.stepContent}>
      <Text size={600} weight="bold">
        What you&#39;ll see
      </Text>

      <div className={styles.featureCardsContainer}>
        {FEATURE_CARDS.map((card) => (
          <Card key={card.emoji}>
            <div className={styles.featureCard}>
              <span className={styles.featureEmoji}>{card.emoji}</span>
              <Text size={300} className={styles.featureDescription}>
                {card.description}
              </Text>
            </div>
          </Card>
        ))}
      </div>

      <div className={styles.buttonRowSpaced}>
        <Button appearance="subtle" onClick={onBack}>
          &larr; Back
        </Button>
        <Button appearance="primary" onClick={onNext}>
          Next →
        </Button>
      </div>
    </div>
  );
}

function Step3Ready({ onBack }: { onBack: () => void }): React.ReactElement {
  const styles = useStyles();

  return (
    <div className={styles.stepContent}>
      <Text size={600} weight="bold">
        You&#39;re all set
      </Text>

      <Text size={300} className={styles.bodyText}>
        The read and compose panes work immediately — no sign-in needed.
      </Text>
      <Text size={300} className={styles.bodyText}>
        To view your Dashboard and full mailbox history, you&#39;ll need to sign in with Microsoft to grant
        read-only metadata access (Mail.ReadBasic scope).
      </Text>

      <div className={styles.buttonRow}>
        <Button
          appearance="primary"
          onClick={() => {
            window.location.search = "?mode=dashboard";
          }}
        >
          Open Dashboard
        </Button>
        <Button
          appearance="secondary"
          onClick={() => {
            window.location.search = "?mode=read";
          }}
        >
          Start reading emails
        </Button>
      </div>

      <div>
        <Button appearance="subtle" onClick={onBack}>
          &larr; Back
        </Button>
      </div>
    </div>
  );
}

// ── Main component ────────────────────────────────────────────────────────────

export default function OnboardingPane(): React.ReactElement {
  const styles = useStyles();
  const [step, setStep] = useState<1 | 2 | 3>(1);

  const goToStep = (s: 1 | 2 | 3) => setStep(s);

  return (
    <div className={styles.root}>
      <Text size={200} className={styles.stepIndicator}>
        Step {step} of 3
      </Text>

      <Divider />

      {step === 1 && (
        <Step1Privacy
          onNext={() => goToStep(2)}
        />
      )}

      {step === 2 && (
        <Step2Features
          onNext={() => goToStep(3)}
          onBack={() => goToStep(1)}
        />
      )}

      {step === 3 && (
        <Step3Ready
          onBack={() => goToStep(2)}
        />
      )}
    </div>
  );
}
