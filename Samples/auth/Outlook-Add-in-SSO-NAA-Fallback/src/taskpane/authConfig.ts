// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console, Office */

import {
  BrowserAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  Configuration,
  LogLevel,
  AuthenticationResult,
} from "@azure/msal-browser";
import { createLocalUrl } from "./util";
import { AccountContext } from "./msalcommon";

export { AccountManager };

const applicationId = "fccd3bcf-08f0-4b8a-b36f-520cfaa4ab51";

function getMsalConfig(enableDebugLogging: boolean) {
  const msalConfig: Configuration = {
    auth: {
      clientId: applicationId,
      authority: "https://login.microsoftonline.com/common",
    },
    system: {},
  };
  if (enableDebugLogging && msalConfig.system) {
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
  private authMethod: string = ""; // Track authentication method (NAA/Dialog/IE MSAL v2)

  // Initialize MSAL public client application.
  async initialize() {
    // Check if running in the IE webview. If so need to use the MSAL v2 library
    if (this.isWebViewIE()) {
      this.initializeForIE();
      return;
    } else {
      this.authMethod = "NAA";

      // If auth is not working, enable debug logging to help diagnose.
      this.pca = await createNestablePublicClientApplication(getMsalConfig(true));
    }
  }

  initializeForIE() {
    this.authMethod = "IE MSAL v2"; //todo
  }

  isWebViewIE() {
    return false; //todo
  }

  /**
   * Gets the user account information object from MSAL.
   *
   * @param scopes the minimum scopes needed.
   * @returns The user account info.
   */
  async ssoGetUserAccount(scopes: string[]) {
    const userAccount = await this.ssoGetAccessToken(scopes);
    return userAccount;
  }

  /**
   *
   * Uses MSAL and nested app authentication to get an access token from Office SSO.
   * This demonstrates how to work with user identity from the token.
   *
   * @param scopes The minimum scopes needed.
   * @returns The access token.
   */
  async ssoGetAccessToken(scopes: string[]) {
    const userAccount = await this.ssoGetUserIdentity(scopes);
    return userAccount.accessToken;
  }

  async getTokenWithDialogApi(isInternetExplorer?: boolean): Promise<string> {
    return new Promise((resolve) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`${isInternetExplorer ? "dialogie.html" : "dialog.html"}`),
        (result) => {
          result.value.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            const parsedMessage = JSON.parse(arg.message);
            resolve(parsedMessage.token);
            result.value.close();
          });
        }
      );
    });
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
    let userAccount: AuthenticationResult | undefined;
    if (this.authMethod === "") {
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
      userAccount = authResult;
    } catch (error) {
      console.log(`Unable to acquire token silently: ${error}`);
    }

    if (userAccount === undefined) {
      // Acquire token silent failure. Send an interactive request via popup.
      try {
        console.log("Trying to acquire token interactively...");
        const authResult = await this.pca.acquireTokenPopup(tokenRequest);
        console.log("Acquired token interactively.");
        userAccount = authResult;
      } catch (popupError) {
        // Optional fallback if about:blank popup should not be shown
        if (popupError instanceof BrowserAuthError && popupError.errorCode === "popup_window_error") {
          let accessToken = await this.getTokenWithDialogApi();
          console.log(accessToken);
        } else {
          // Acquire token interactive failure.
          console.log(`Unable to acquire token interactively: ${popupError}`);
          throw new Error(`Unable to acquire access token: ${popupError}`);
        }
      }
    }
    return userAccount;
  }
}
