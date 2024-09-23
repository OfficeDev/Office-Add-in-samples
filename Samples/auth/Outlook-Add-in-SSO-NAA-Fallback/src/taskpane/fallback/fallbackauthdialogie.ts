/* global Office */

import { PublicClientApplication } from "@azure/msal-browser-v2";
import { defaultScopes, getMsalConfig } from "../msalconfig";
import { createLocalUrl } from "../util";
import { UserProfile } from "../authhelper";

export async function initializeMsal() {
  const publicClientApp = new PublicClientApplication(getMsalConfig(true));
  try {
    const result = await publicClientApp.handleRedirectPromise();
    if (result) {
      publicClientApp.setActiveAccount(result.account);
      await Office.onReady();
      const idTokenClaims = result.idTokenClaims as { name?: string; preferred_username?: string };
      const userProfile: UserProfile = {
        userName: idTokenClaims.name,
        userEmail: idTokenClaims.preferred_username,
        accessToken: result.accessToken,
      };
      Office.context.ui.messageParent(JSON.stringify(userProfile));
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
