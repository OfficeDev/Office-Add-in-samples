/* global Office */

import { PublicClientApplication } from "@azure/msal-browser-v2";
import { defaultScopes, msalConfig } from "./msalconfig";
import { createLocalUrl } from "./util";

export async function initializeMsal() {
  const publicClientApp = new PublicClientApplication(msalConfig);
  try {
    const result = await publicClientApp.handleRedirectPromise();
    if (result) {
      publicClientApp.setActiveAccount(result.account);
      await Office.onReady();
      Office.context.ui.messageParent(JSON.stringify({ token: result.accessToken }));
      return;
    }
  } catch (ex) {
    await Office.onReady();
    Office.context.ui.messageParent(JSON.stringify({ error: ex.name }));
    return;
  }

  publicClientApp.acquireTokenRedirect({
    scopes: defaultScopes,
    redirectUri: createLocalUrl("dialogie.html"),
  });
}
initializeMsal();
