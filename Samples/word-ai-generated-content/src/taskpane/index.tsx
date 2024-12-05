/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import Home from "./components/Home";

/* global document, Office, module, require */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady().then(() => {
  root.render(
    <FluentProvider theme={webLightTheme}>
      <Home />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/Home", () => {
    const NextApp = require("./components/Home").default;
    root.render(NextApp);
  });
}
