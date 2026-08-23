import React, { useState, useEffect } from "react";
import { FluentProvider, webLightTheme, Spinner } from "@fluentui/react-components";
import ReadPane from "./ReadPane";
import ComposePane from "./ComposePane";
import DashboardPane from "./DashboardPane";
import SettingsPane from "./SettingsPane";
import MethodologyPane from "./MethodologyPane";
import OnboardingPane from "./OnboardingPane";

function getMode(): string | null {
  const params = new URLSearchParams(window.location.search);
  return params.get("mode");
}

export default function App(): React.ReactElement {
  const mode = getMode();

  // When no ?mode= param is present, we need to detect the context from Office.js.
  // null = still detecting, "compose" | "read" | "onboarding" = resolved.
  const [detectedMode, setDetectedMode] = useState<string | null>(
    mode !== null ? mode : null
  );
  const [detecting, setDetecting] = useState<boolean>(mode === null);

  useEffect(() => {
    // Only run detection when no explicit ?mode= param was supplied.
    if (mode !== null) return;

    // Office.onReady ensures Office.js has initialised before we inspect context.
    if (typeof Office === "undefined") {
      // Running in a plain browser (e.g. dev preview) — fall back to onboarding.
      setDetectedMode("onboarding");
      setDetecting(false);
      return;
    }

    Office.onReady(() => {
      try {
        const item = Office.context?.mailbox?.item;
        if (!item) {
          // Office is available but no item is loaded — show onboarding.
          setDetectedMode("onboarding");
        } else if (item.itemId) {
          // itemId is a non-empty string only in read mode.
          setDetectedMode("read");
        } else {
          // itemId is undefined/null in compose mode.
          setDetectedMode("compose");
        }
      } catch {
        // Any unexpected error — fall back to onboarding.
        setDetectedMode("onboarding");
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

  return (
    <FluentProvider theme={webLightTheme}>
      {content}
    </FluentProvider>
  );
}
