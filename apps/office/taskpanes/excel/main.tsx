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
    <ExcelTaskpaneApp
      devProfile={{
        label: import.meta.env.MILTON_PUBLIC_DEV_LABEL ?? "",
      }}
      openAI={{
        apiKey: import.meta.env.DEBUG_OPENAI_API_KEY ?? "",
        model: import.meta.env.DEBUG_OPENAI_MODEL ?? "gpt-5-mini",
      }}
    />
  </StrictMode>
);
