import React, { useState, useEffect } from "react";
import { FluentProvider, webLightTheme, Spinner, makeStyles, tokens, Text } from "@fluentui/react-components";
import ReadPane from "./ReadPane";
import ComposePane from "./ComposePane";
import DashboardPane from "./DashboardPane";
import SettingsPane from "./SettingsPane";
import MethodologyPane from "./MethodologyPane";
import OnboardingPane from "./OnboardingPane";

// ── Nav bar styles ────────────────────────────────────────────────────────────

const useNavStyles = makeStyles({
  navBar: {
    position: "fixed",
    bottom: "0",
    left: "0",
    right: "0",
    background: tokens.colorNeutralBackground1,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    zIndex: "100",
    display: "flex",
    justifyContent: "space-around",
    padding: "6px 0",
  },
  navTab: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "2px",
    padding: "4px 8px",
    border: "none",
    background: "transparent",
    cursor: "pointer",
    borderRadius: tokens.borderRadiusMedium,
    minWidth: "60px",
    ":hover": {
      background: tokens.colorNeutralBackground1Hover,
    },
  },
  navTabActive: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "2px",
    padding: "4px 8px",
    border: "none",
    background: tokens.colorBrandBackground2,
    cursor: "pointer",
    borderRadius: tokens.borderRadiusMedium,
    minWidth: "60px",
  },
  navIcon: {
    fontSize: "18px",
    lineHeight: "1",
  },
  navLabel: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
    lineHeight: "1",
  },
  navLabelActive: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: "1",
  },
  contentWrapper: {
    paddingBottom: "60px",
  },
});

function getMode(): string | null {
  const params = new URLSearchParams(window.location.search);
  return params.get("mode");
}

const NAV_TABS = [
  { mode: "read",      icon: "📧", label: "Read" },
  { mode: "compose",   icon: "✏️", label: "Compose" },
  { mode: "dashboard", icon: "📊", label: "Dashboard" },
  { mode: "settings",  icon: "⚙️", label: "Settings" },
] as const;

export default function App(): React.ReactElement {
  const navStyles = useNavStyles();
  const mode = getMode();

  // When no ?mode= param is present, we need to detect the context from Office.js.
  // null = still detecting, "compose" | "read" | "onboarding" = resolved.
  const [detectedMode, setDetectedMode] = useState<string | null>(null);
  // Always wait for Office.onReady so panes never mount before Office initialises.
  const [detecting, setDetecting] = useState<boolean>(true);

  useEffect(() => {
    if (typeof Office === "undefined") {
      if (mode === null) setDetectedMode("onboarding");
      setDetecting(false);
      return;
    }

    Office.onReady(() => {
      if (mode === null) {
        try {
          const item = Office.context?.mailbox?.item;
          if (!item) {
            setDetectedMode("onboarding");
          } else if (item.itemId) {
            setDetectedMode("read");
          } else {
            setDetectedMode("compose");
          }
        } catch {
          setDetectedMode("onboarding");
        }
      }
      setDetecting(false);
    });
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Determine which mode string to render from.
  const resolvedMode = mode !== null ? mode : detectedMode;

  let content: React.ReactElement;

  if (detecting) {
    // Show a brief spinner while Office.onReady fires and we detect the context.
    content = (
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", height: "100vh" }}>
        <Spinner label="Loading…" size="medium" />
      </div>
    );
  } else if (resolvedMode === null || resolvedMode === "onboarding") {
    content = <OnboardingPane />;
  } else if (resolvedMode === "read") {
    content = <ReadPane />;
  } else if (resolvedMode === "compose") {
    content = <ComposePane />;
  } else if (resolvedMode === "dashboard") {
    content = <DashboardPane />;
  } else if (resolvedMode === "settings") {
    content = <SettingsPane />;
  } else if (resolvedMode === "methodology") {
    content = <MethodologyPane />;
  } else {
    content = (
      <div style={{ padding: "16px" }}>
        <p>
          Unknown mode: &quot;{resolvedMode}&quot;. Expected &quot;read&quot;, &quot;compose&quot;,
          &quot;dashboard&quot;, &quot;settings&quot;, &quot;methodology&quot;, or &quot;onboarding&quot;.
        </p>
      </div>
    );
  }

  const showNav =
    !detecting &&
    resolvedMode !== null &&
    resolvedMode !== "onboarding" &&
    resolvedMode !== "methodology";

  return (
    <FluentProvider theme={webLightTheme}>
      <div className={showNav ? navStyles.contentWrapper : undefined}>
        {content}
      </div>
      {showNav && (
        <nav className={navStyles.navBar} aria-label="Main navigation">
          {NAV_TABS.map(({ mode: tabMode, icon, label }) => {
            const isActive = resolvedMode === tabMode;
            return (
              <button
                key={tabMode}
                className={isActive ? navStyles.navTabActive : navStyles.navTab}
                onClick={() => { window.location.search = `?mode=${tabMode}`; }}
                aria-current={isActive ? "page" : undefined}
              >
                <span className={navStyles.navIcon}>{icon}</span>
                <Text className={isActive ? navStyles.navLabelActive : navStyles.navLabel}>
                  {label}
                </Text>
              </button>
            );
          })}
        </nav>
      )}
    </FluentProvider>
  );
}
