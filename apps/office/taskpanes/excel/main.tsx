import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { ExcelTaskpaneApp } from "@milton/ui";
import "@milton/ui/styles.css";

const rootElement = document.getElementById("root");

if (!rootElement) {
  throw new Error("Missing root element.");
}

const root = createRoot(rootElement);

root.render(
  <StrictMode>
    <ExcelTaskpaneApp />
  </StrictMode>
);
