import * as React from "react";
import { createRoot } from "react-dom/client";
// import App from "../taskpane/components/App";
// import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import UserInputDialog from "./components/UserInputDialog";

/* global document, Office, module, require, HTMLElement console */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  root?.render(
    <UserInputDialog
      isOpen={false}
      onClose={function (): void {
        throw new Error("Function not implemented.");
      }}
      onSubmit={function (input: string): void {
        // throw new Error("Function not implemented.");
        console.log(input);
      }}
    />
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("../taskpane/components/App", () => {
    const NextApp = require("../taskpane/components/App").default;
    root?.render(NextApp);
  });
}
