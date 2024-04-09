// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console */

import { PublicClientNext, type IPublicClientApplication } from "@azure/msal-browser";

export { AccountManager };

const applicationId = "Enter_the_Application_Id_Here";

const msalConfig = {
  auth: {
    clientId: applicationId,
    authority: "https://login.microsoftonline.com/common",
    supportsNestedAppAuth: true,
  },
};

// Encapsulate functions for getting user account and token information.
class AccountManager {
  loginHint: string = "";
  pca: IPublicClientApplication | undefined = undefined;

  // Initialize MSAL public client application.
  async initialize(loginHint: string) {
    this.loginHint = loginHint;
    this.pca = await PublicClientNext.createPublicClientApplication(msalConfig);
  }

  /**
   * 
   * @param scopes the minimum scopes needed.
   * @returns An access token.
   */
  async ssoGetToken(scopes: string[]) {
    const userAccount = await this.ssoGetUserIdentity(scopes);
    return userAccount.accessToken;
  }

  /**
   * 
   * Uses MSAL and nested app authentication to get the user account from Office SSO.
   * This demonstrates how to work with user identity from the token.
   * 
   * @param scopes The minimum scopes needed.
   * @returns The user account data (including identity).
   */
  async ssoGetUserIdentity(scopes: string[]) {
    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }

    // Specify minimum scopes needed for the access token.
    const tokenRequest = {
      scopes: scopes,
      loginHint: this.loginHint,
    };

    try {
      console.log("Trying to acquire token silently...");

      //acquireTokenSilent requires an active account. Check if one exists, otherwise use ssoSilent.
      const account = this.pca.getActiveAccount();
      const userAccount = account ? await this.pca.acquireTokenSilent(tokenRequest) : await this.pca.ssoSilent(tokenRequest);

      console.log("Acquired token silently.");
      return userAccount;
    } catch (error) {
      console.log(`Unable to acquire token silently: ${error}`);
    }

    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const userAccount = await this.pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      return userAccount;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Unable to acquire token interactively: ${popupError}`);
      throw new Error(`Unable to acquire access token: ${popupError}`);
    }
  }
}
