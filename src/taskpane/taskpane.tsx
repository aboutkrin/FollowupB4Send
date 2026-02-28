import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { ReminderForm } from "./components/ReminderForm";

Office.onReady(() => {
  const container = document.getElementById("root")!;
  const root = createRoot(container);
  root.render(
    <FluentProvider theme={webLightTheme}>
      <ReminderForm />
    </FluentProvider>
  );
});
