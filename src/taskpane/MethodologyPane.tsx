import React from "react";
import {
  Text,
  Card,
  Divider,
  Link,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { METHODOLOGY_VERSION } from "../engine/calcEngine";

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
    padding: tokens.spacingHorizontalL,
  },
  header: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  subtitle: {
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
  },
  formulaBox: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontFamily: "monospace",
    fontSize: tokens.fontSizeBase300,
    overflowX: "auto",
  },
  table: {
    width: "100%",
    borderCollapse: "collapse" as const,
    fontSize: tokens.fontSizeBase200,
  },
  th: {
    textAlign: "left" as const,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
  },
  td: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: "top" as const,
  },
  tdMono: {
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    fontFamily: "monospace",
    verticalAlign: "top" as const,
  },
  bulletList: {
    margin: 0,
    paddingLeft: "1.25rem",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  bulletItem: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  equivalencyCard: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingVerticalS,
  },
  sourceItem: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    paddingBottom: tokens.spacingVerticalXS,
  },
  footer: {
    display: "flex",
    justifyContent: "flex-start",
    paddingTop: tokens.spacingVerticalS,
  },
});

// ── Component ─────────────────────────────────────────────────────────────────

export default function MethodologyPane(): React.ReactElement {
  const styles = useStyles();

  return (
    <div className={styles.root}>

      {/* 1. Header */}
      <div className={styles.header}>
        <Text size={600} weight="bold">
          How eCFP calculates your footprint
        </Text>
        <Text size={200} className={styles.subtitle}>
          Methodology v{METHODOLOGY_VERSION} · All computation client-side
        </Text>
      </div>

      <Divider />

      {/* 2. Send/Receive formula */}
      <div className={styles.section}>
        <Text size={400} weight="semibold" className={styles.sectionTitle}>
          Send / Receive formula
        </Text>
        <Text size={200} color="brand">
          §4.3 of the specification
        </Text>
        <div className={styles.formulaBox}>
          gCO₂e = (E_base + E_data × size_MB × I_mix) × R_factor
        </div>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th}>Constant</th>
              <th className={styles.th}>Value</th>
              <th className={styles.th}>Meaning</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className={styles.tdMono}>E_base</td>
              <td className={styles.tdMono}>0.3 g</td>
              <td className={styles.td}>
                Short laptop email baseline (device + network + server + amortised embodied carbon)
              </td>
            </tr>
            <tr>
              <td className={styles.tdMono}>E_data</td>
              <td className={styles.tdMono}>11 g/MB</td>
              <td className={styles.td}>
                Transmission and processing coefficient
              </td>
            </tr>
            <tr>
              <td className={styles.tdMono}>I_mix</td>
              <td className={styles.tdMono}>1.0 / 0.6</td>
              <td className={styles.td}>
                1.0 for internet recipients; 0.6 when all recipients share sender&#39;s domain (intranet discount)
              </td>
            </tr>
            <tr>
              <td className={styles.tdMono}>R_factor</td>
              <td className={styles.tdMono}>1 + 0.25×(N−1), max 3.0</td>
              <td className={styles.td}>
                Scales footprint with recipient count
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <Divider />

      {/* 3. Storage formula */}
      <div className={styles.section}>
        <Text size={400} weight="semibold" className={styles.sectionTitle}>
          Storage formula
        </Text>
        <Text size={200} color="brand">
          §4.4 of the specification
        </Text>
        <div className={styles.formulaBox}>
          gCO₂e_storage = size_MB × age_years × E_store
        </div>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th}>Constant</th>
              <th className={styles.th}>Value</th>
              <th className={styles.th}>Meaning</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className={styles.tdMono}>E_store</td>
              <td className={styles.tdMono}>10 g/MB/year</td>
              <td className={styles.td}>
                Data-center storage energy coefficient (placeholder — literature review pending)
              </td>
            </tr>
          </tbody>
        </table>
        <Text size={200} color="neutral">
          Long-term email retention contributes ongoing storage energy. Deleting old mail reduces this.
        </Text>
      </div>

      <Divider />

      {/* 4. Known limitations */}
      <div className={styles.section}>
        <Text size={400} weight="semibold" className={styles.sectionTitle}>
          Known limitations
        </Text>
        <ul className={styles.bulletList}>
          <li className={styles.bulletItem}>
            Size-driven model: long plain-text emails may be underestimated vs. device-time-driven approaches
            (published ~17 g figure reflects time spent writing/reading, not data size)
          </li>
          <li className={styles.bulletItem}>
            Network hops (factor #4) not observable client-side — folded into published averages
          </li>
          <li className={styles.bulletItem}>
            Device power (factor #7) represented via published coefficients, not live telemetry
          </li>
          <li className={styles.bulletItem}>
            E_store coefficient is a placeholder pending literature review
          </li>
        </ul>
      </div>

      <Divider />

      {/* 5. Calibration */}
      <div className={styles.section}>
        <Text size={400} weight="semibold" className={styles.sectionTitle}>
          Calibration
        </Text>
        <Text size={200} color="brand">
          §4.5 of the specification
        </Text>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th}>Email type</th>
              <th className={styles.th}>Published</th>
              <th className={styles.th}>This model</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className={styles.td}>Short reply</td>
              <td className={styles.tdMono}>~0.3 g</td>
              <td className={styles.tdMono}>~0.3 g</td>
            </tr>
            <tr>
              <td className={styles.td}>Image / attachment email</td>
              <td className={styles.tdMono}>up to ~50 g</td>
              <td className={styles.tdMono}>11–50 g (size-dependent)</td>
            </tr>
            <tr>
              <td className={styles.td}>Office worker annual</td>
              <td className={styles.tdMono}>~184 kg/yr</td>
              <td className={styles.td}>Dashboard sanity check</td>
            </tr>
          </tbody>
        </table>
      </div>

      <Divider />

      {/* 6. Equivalencies */}
      <div className={styles.section}>
        <Text size={400} weight="semibold" className={styles.sectionTitle}>
          Real-world equivalencies
        </Text>
        <Text size={200} color="brand">
          §4.6 of the specification
        </Text>
        <Card>
          <div className={styles.equivalencyCard}>
            <Text size={300}>
              🚗 1 km driven (avg gasoline car) ≈ 170 g CO₂e
            </Text>
            <Text size={300}>
              🔋 1 smartphone charge ≈ 8 g CO₂e
            </Text>
          </div>
        </Card>
      </div>

      <Divider />

      {/* 7. Sources */}
      <div className={styles.section}>
        <Text size={400} weight="semibold" className={styles.sectionTitle}>
          Sources
        </Text>
        <Text size={200} color="brand">
          Appendix A of the specification
        </Text>
        <ol className={styles.bulletList}>
          <li className={styles.sourceItem}>
            Berners-Lee, M. <em>How Bad Are Bananas?</em> (2020 ed.) — underlying per-email methodology
          </li>
          <li className={styles.sourceItem}>
            AgainstData, &#34;Email Carbon Footprint&#34; (Jun 2025) — per-email figures, office-worker benchmark, storage-vs-deletion argument
          </li>
          <li className={styles.sourceItem}>
            The Shift Project, <em>Lean ICT</em> (2019) — supporting coefficients
          </li>
          <li className={styles.sourceItem}>
            Carbon Literacy Project — supporting coefficients
          </li>
        </ol>
      </div>

      <Divider />

      {/* 8. Footer */}
      <div className={styles.footer}>
        <Link
          onClick={() => {
            window.location.search = "?mode=read";
          }}
        >
          &larr; Back
        </Link>
      </div>

    </div>
  );
}
