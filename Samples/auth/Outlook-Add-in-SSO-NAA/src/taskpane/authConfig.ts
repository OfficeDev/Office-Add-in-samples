// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console*/

import {
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  Configuration,
  LogLevel,
} from "@azure/msal-browser";

export { AccountManager };

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
    if (msalConfig.system) {
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
  }
  return msalConfig;
}

// Encapsulate functions for getting user account and token information.
class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;

  // Initialize MSAL public client application.
  async initialize() {
    // If auth is not working, enable debug logging to help diagnose.
    this.pca = await createNestablePublicClientApplication(getMsalConfig(false));
  }

  /**
   *
   * @param scopes the minimum scopes needed.
   * @returns An access token.
   */
  async ssoGetAccessToken(scopes: string[]) {
    const userAccount = await this.ssoGetUserAccount(scopes);
    return userAccount.accessToken;
  }

  /**
   *
   * Uses MSAL and nested app authentication to get the user account from Office SSO.
   *
   * @param scopes The minimum scopes needed.
   * @returns The user account information from MSAL.
   */
  async ssoGetUserAccount(scopes: string[]) {
    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }

    // Specify minimum scopes needed for the access token.
    const tokenRequest = {
      scopes: scopes,
    };

    try {
      console.log("Trying to acquire token silently...");
      const authResult = await this.pca.acquireTokenSilent(tokenRequest);
      console.log("Acquired token silently.");
      return authResult;
    } catch (error) {
      console.log(`Unable to acquire token silently: ${error}`);
    }

    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const authResult = await this.pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      return authResult;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Unable to acquire token interactively: ${popupError}`);
      throw new Error(`Unable to acquire access token: ${popupError}`);
    }
  }
}
