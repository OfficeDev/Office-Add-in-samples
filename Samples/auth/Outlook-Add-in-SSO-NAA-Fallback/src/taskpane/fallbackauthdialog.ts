/* global Office, window */

import { AuthenticationResult } from "@azure/msal-browser";
import { getTokenRequest, AccountContext, ensurePublicClient } from "./msalcommon";
import { createLocalUrl } from "./util";
import { PublicClientApplication } from "@azure/msal-browser";
import { UserProfile } from "./userProfile";

// read querystring parameter
function getQueryParameter(param: string) {
  const params = new URLSearchParams(window.location.search);
  return params.get(param);
}

async function returnResult(publicClientApp: PublicClientApplication, authResult: AuthenticationResult) {
  publicClientApp.setActiveAccount(authResult.account);
  await Office.onReady();
  const idTokenClaims = authResult.idTokenClaims as { name?: string; preferred_username?: string };
  const userProfile: UserProfile = {
    userName: idTokenClaims.name,
    userEmail: idTokenClaims.preferred_username,
    accessToken: authResult.accessToken,
  };
  Office.context.ui.messageParent(JSON.stringify(userProfile));
  return;
}
export async function initializeMsal() {
  const publicClientApp = await ensurePublicClient();
  try {
    if (getQueryParameter("logout") === "1") {
      await publicClientApp.logoutRedirect();
      return;
    }
    const result = await publicClientApp.handleRedirectPromise();

    if (result) {
      return returnResult(publicClientApp, result);
    }
  } catch (ex) {
    await Office.onReady();
    Office.context.ui.messageParent(JSON.stringify({ error: ex.name }));
    return;
  }

  const accountContextString = getQueryParameter("accountContext");
  let accountContext: AccountContext;
  if (accountContextString) {
    accountContext = JSON.parse(accountContextString);
  }
  const request = await getTokenRequest(accountContext);
  try {
    publicClientApp.acquireTokenRedirect({
      ...request,
      redirectUri: createLocalUrl("dialog.html"),
    });
  } catch (ex) {
    const result = await publicClientApp.ssoSilent(request);
    return returnResult(publicClientApp, result);
  }
}

initializeMsal();
