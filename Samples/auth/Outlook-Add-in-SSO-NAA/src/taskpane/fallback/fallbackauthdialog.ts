// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file handls MSAL auth for the fallback dialog page. */

/* global Office, window */

import { AuthenticationResult } from "@azure/msal-browser";
import { getTokenRequest, AccountContext, ensurePublicClient } from "../msalcommon";
import { createLocalUrl } from "../util";
import { PublicClientApplication } from "@azure/msal-browser";

// read querystring parameter
function getQueryParameter(param: string) {
  const params = new URLSearchParams(window.location.search);
  return params.get(param);
}

async function returnResult(publicClientApp: PublicClientApplication, authResult: AuthenticationResult) {
  publicClientApp.setActiveAccount(authResult.account);
  await Office.onReady();
  Office.context.ui.messageParent(JSON.stringify({ token: authResult.accessToken }));
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
