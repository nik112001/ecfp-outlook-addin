import React from "react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
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

  let content: React.ReactElement;
  if (mode === null) {
    // No ?mode= param: show onboarding for first-time users
    content = <OnboardingPane />;
  } else if (mode === "read") {
    content = <ReadPane />;
  } else if (mode === "compose") {
    content = <ComposePane />;
  } else if (mode === "dashboard") {
    content = <DashboardPane />;
  } else if (mode === "settings") {
    content = <SettingsPane />;
  } else if (mode === "methodology") {
    content = <MethodologyPane />;
  } else if (mode === "onboarding") {
    content = <OnboardingPane />;
  } else {
    content = (
      <div style={{ padding: "16px" }}>
        <p>
          Unknown mode: &quot;{mode}&quot;. Expected &quot;read&quot;, &quot;compose&quot;,
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
