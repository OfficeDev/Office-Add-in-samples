// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file handles the sign out dialog for MSAL V2 (IE11 Trident webview).

import { PublicClientApplication } from "@azure/msal-browser-v2";
import { getMsalConfig } from "../msalConfigV2";
import { shouldCloseDialog, sendDialogMessage, getCurrentPageUrl } from "../util";

export async function initializeMsal() {
  if (shouldCloseDialog()) {
    sendDialogMessage(JSON.stringify({ status: "success" }));
    return;
  }

  try {
    const publicClientApp = new PublicClientApplication(getMsalConfig(true));
    publicClientApp.logoutRedirect({ postLogoutRedirectUri: getCurrentPageUrl({ close: "1" }) });
  } catch (ex: any) {
    sendDialogMessage(JSON.stringify({ error: ex.name }));
  }
}
initializeMsal();
