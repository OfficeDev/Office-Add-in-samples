// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file handles the sign out dialog for MSAL V3.

/* global Office */

import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "../msalconfig";

export async function initializeMsal() {
  const publicClientApp = new PublicClientApplication(msalConfig);
  try {
    await publicClientApp.logoutRedirect();
    await Office.onReady();
    Office.context.ui.messageParent(JSON.stringify({ status: "success" }));
    return;
  } catch (ex) {
    await Office.onReady();
    Office.context.ui.messageParent(JSON.stringify({ error: ex.name }));
    return;
  }
}
initializeMsal();
