import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

function renderApp(itemIdProp: string | undefined) {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App item={itemIdProp ?? ""} />
    </FluentProvider>
  );
}

/* Render application after Office initializes (guard Office for non-host environments) */
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    let itemId: string | undefined = undefined;
    try {
      itemId = Office?.context?.mailbox?.item?.itemId;
    } catch (e) {
      // swallow errors when Office context isn't available yet
      itemId = undefined;
    }
    renderApp(itemId);
  });
} else {
  // If Office is not present (e.g., during local dev outside host), render with empty item
  renderApp(undefined);
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    // Recompute itemId when hot-reloading in a host where Office is available
    let itemId: string | undefined = undefined;
    try {
      itemId = typeof Office !== "undefined" ? Office?.context?.mailbox?.item?.itemId : undefined;
    } catch (e) {
      itemId = undefined;
    }
    const NextApp = require("./components/App").default;
    renderApp(itemId);
  });
}
