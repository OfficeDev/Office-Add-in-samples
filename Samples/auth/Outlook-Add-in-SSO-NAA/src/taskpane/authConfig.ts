// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console, Office */

import {
  type AccountInfo,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  Configuration,
  LogLevel,
} from "@azure/msal-browser";

export { AccountManager };

interface AuthContext {
  loginHint: string;
  userPrincipalName: string;
  userObjectId: string;
  tenantId: string;
}

const applicationId = "Enter_the_Application_Id_Here";

function getMsalConfig(enableDebugLogging: boolean) {
  const msalConfig: Configuration = {
    auth: {
      clientId: applicationId,
      authority: "https://login.microsoftonline.com/common",
    },
    system: {},
  };
  if (enableDebugLogging) {
    msalConfig.system.loggerOptions = {
      logLevel: LogLevel.Verbose,
      loggerCallback: (level: LogLevel, message: string) => {
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
      piiLoggingEnabled: true,
    };
  }
  return msalConfig;
}

// Encapsulate functions for getting user account and token information.
class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private account: AccountInfo | undefined = undefined;
  private loginHint: string | undefined = undefined;

  // Initialize MSAL public client application.
  async initialize() {
    // If auth is not working, enable debug logging to help diagnose.
    this.pca = await createNestablePublicClientApplication(getMsalConfig(false));

    // Initialize account by matching account known by Outlook with MSAL.js
    try {
      const authContext: AuthContext = await Office.auth.getAuthContext();
      const username = authContext.userPrincipalName;
      const tenantId = authContext.tenantId;
      const localAccountId = authContext.userObjectId;
      this.loginHint = authContext.loginHint || authContext.userPrincipalName;
      const account = this.pca.getAccount({
        username,
        localAccountId,
        tenantId,
      });
      if (account) {
        this.account = account;
      }
    } catch {
      // Intentionally empty catch block.
    }

    if (!this.loginHint) {
      const accountType = Office.context.mailbox.userProfile.accountType;
      this.loginHint =
        accountType === "office365" || accountType === "outlookCom"
          ? Office.context.mailbox.userProfile.emailAddress
          : "";
    }
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
      account: this.account,
    };

    try {
      console.log("Trying to acquire token silently...");

      //acquireTokenSilent requires an active account. Check if one exists, otherwise use ssoSilent.
      const authResult = this.account
        ? await this.pca.acquireTokenSilent(tokenRequest)
        : await this.pca.ssoSilent(tokenRequest);
      this.account = authResult.account;

      console.log("Acquired token silently.");
      return authResult;
    } catch (error) {
      console.log(`Unable to acquire token silently: ${error}`);
    }

    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const authResult = await this.pca.acquireTokenPopup(tokenRequest);
      this.account = authResult.account;
      console.log("Acquired token interactively.");
      return authResult;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Unable to acquire token interactively: ${popupError}`);
      throw new Error(`Unable to acquire access token: ${popupError}`);
    }
  }
}
