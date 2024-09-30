// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console, Office */

import {
  BrowserAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { getMsalConfig } from "./msalConfigV3";
import { FallbackError } from "./errors";

// Encapsulate functions for getting user account and token information.
export class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;

  // Initialize MSAL public client application.
  async initialize() {
    // If auth is not working, enable debug logging to help diagnose.
    this.pca = await createNestablePublicClientApplication(getMsalConfig(true));
  }

  public isNestedAppAuthSupported() {
    return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  }

  public hasActiveAccount() {
    return this.pca?.getActiveAccount() ? true : false;
  }

  /**
   *
   * Uses MSAL and nested app authentication to get an access token from Office SSO.
   * This demonstrates how to work with user identity from the token.
   *
   * @param scopes The minimum scopes needed.
   * @returns The access token.
   */
  public async ssoGetAccessToken(scopes: string[]): Promise<string> {
    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }

    const tokenRequest = {
      scopes,
    };
    try {
      console.log("Trying to acquire token silently...");
      const authResult = await this.pca.acquireTokenSilent(tokenRequest);
      console.log("Acquired token silently.");
      return authResult.accessToken;
    } catch (error) {
      console.warn(`Unable to acquire token silently: ${error}`);
    }

    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const selectAccount = this.pca.getActiveAccount() ? false : true;
      const interactiveRequest = {
        ...tokenRequest,
        ...(selectAccount ? { prompt: "select_account" } : {}),
      };

      const authResult = await this.pca.acquireTokenPopup(interactiveRequest);
      console.log("Acquired token interactively.");
      if (selectAccount) {
        this.pca.setActiveAccount(authResult.account);
      }

      return authResult.accessToken;
    } catch (popupError) {
      if (popupError instanceof BrowserAuthError && popupError.errorCode === "popup_window_error") {
        throw new FallbackError("msal-browser fallback not supported");
      }
      throw popupError;
    }
  }

  public async signOut() {
    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }
    await this.pca.logoutPopup();
  }
}
