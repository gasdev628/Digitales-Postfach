import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { createLightTheme, FluentProvider } from "@fluentui/react-components";
import type { BrandVariants, Theme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement */

// themes designer: https://react.fluentui.dev/?path=/docs/theme-theme-designer--docs
const outlook: BrandVariants = {
  10: "#010101",
  20: "#010202",
  30: "#010403",
  40: "#0D1C17",
  50: "#0F2F25",
  60: "#0E3D30",
  70: "#094B3A",
  80: "#005B46",
  90: "#216753",
  100: "#377461",
  110: "#4B816F",
  120: "#5F8E7D",
  130: "#729B8C",
  140: "#85A99C",
  150: "#98B7AB",
  160: "#ACC4BB",
};

const customTheme: Theme = {
  ...createLightTheme(outlook),
};

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  root?.render(
    <FluentProvider theme={customTheme}>
      <App />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
