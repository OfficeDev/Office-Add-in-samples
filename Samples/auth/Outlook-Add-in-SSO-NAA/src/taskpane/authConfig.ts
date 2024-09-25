// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global Office, console, window*/

import {
  BrowserAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { msalConfig } from "./msalconfig";

export { AccountManager };

type AccountContext = {
  loginHint?: string;
  tenantId?: string;
  localAccountId?: string;
};

// Encapsulate functions for getting user account and token information.
class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private _authContext: AccountContext | null = null;

  // Initialize MSAL public client application.
  async initialize() {
    // If auth is not working, enable debug logging to help diagnose.
    this.pca = await createNestablePublicClientApplication(msalConfig);
  }

  /**
   *
   * @param scopes the minimum scopes needed.
   * @returns An access token.
   */
  async ssoGetAccessToken(scopes: string[]) {
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
      return authResult.accessToken;
    } catch (error) {
      console.log(`Unable to acquire token silently: ${error}`);
    }

    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const authResult = await this.pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      return authResult.accessToken;
    } catch (popupError) {
      // Optional fallback if about:blank popup should not be shown
      if (popupError instanceof BrowserAuthError && popupError.errorCode === "popup_window_error") {
        const accessToken = await this.getTokenWithDialogApi();
        return accessToken;
      } else {
        // Acquire token interactive failure.
        console.log(`Unable to acquire token interactively: ${popupError}`);
        throw new Error(`Unable to acquire access token: ${popupError}`);
      }
    }
  }

  async getTokenWithDialogApi(): Promise<string> {
    const accountContext = await this.getAccountContext();
    return new Promise((resolve) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html?accountContext=${encodeURIComponent(JSON.stringify(accountContext))}`),
        { height: 60, width: 30 },
        (result) => {
          result.value.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (arg: { message: string; origin: string | undefined }) => {
              const parsedMessage = JSON.parse(arg.message);
              resolve(parsedMessage.token);
              result.value.close();
            }
          );
        }
      );
    });
  }
  async getAccountContext(): Promise<AccountContext | null> {
    if (!this._authContext) {
      try {
        const authContext = await (Office.auth as any).getAuthContext();
        this._authContext = {
          loginHint: authContext.loginHint,
          tenantId: authContext.tenantId,
          localAccountId: authContext.userObjectId,
        };
      } catch {
        this._authContext = {};
      }
    }
    return this._authContext;
  }
}

function createLocalUrl(path: string) {
  return `${window.location.origin}/${path}`;
}
