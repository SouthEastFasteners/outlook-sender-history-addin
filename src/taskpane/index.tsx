import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import "./taskpane.css";

/* global Office */
Office.onReady(() => {
  const container = document.getElementById("root")!;
  const root = createRoot(container);
  root.render(<App />);
});
