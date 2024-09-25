// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global Office, console*/

import {
  BrowserAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { msalConfig } from "./msalconfig";
import { getAccountFromContext } from "./msalcommon";
import { createLocalUrl, setSignOutButtonVisibility } from "./util";

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

  /**
   * Gets an access token by using the Office dialog API to handle authentication. Used for fallback scenario.
   * @returns The access token.
   */
  async getTokenWithDialogApi(): Promise<string> {
    const accountContext = await getAccountFromContext();
    return new Promise((resolve) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html?accountContext=${encodeURIComponent(JSON.stringify(accountContext))}`),
        { height: 60, width: 30 },
        (result) => {
          result.value.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (arg: { message: string; origin: string | undefined }) => {
              const parsedMessage = JSON.parse(arg.message);
              result.value.close();
              setSignOutButtonVisibility(true);
              resolve(parsedMessage.token);
            }
          );
        }
      );
    });
  }
}
