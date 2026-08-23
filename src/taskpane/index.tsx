import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Office: any;

Office.onReady(() => {
  const container = document.getElementById("root");
  if (!container) {
    throw new Error("Could not find #root element");
  }
  const root = createRoot(container);
  root.render(<App />);
});
