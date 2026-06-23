import * as React from "react";
import { createRoot } from "react-dom/client";
import ColorPickerApp from "./ColorPickerApp";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

const title = "ExcelStarch";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <ColorPickerApp title={title} />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./ColorPickerApp", () => {
    const NextApp = require("./ColorPickerApp").default;
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <NextApp title={title} />
      </FluentProvider>
    );
  });
}
