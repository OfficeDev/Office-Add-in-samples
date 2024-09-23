// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console, Office */

import {
  BrowserAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  AuthenticationResult,
} from "@azure/msal-browser";
import { createLocalUrl } from "./util";
import { getMsalConfig } from "./msalconfig";
import { UserProfile } from "./userProfile";

export { AccountManager };

// Encapsulate functions for getting user account and token information.
class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private fallbackPopup = false; // true if Outlook client has the about:blank popup bug and we need to fall back.
  private gUserProfile: UserProfile = {};

  // Initialize MSAL public client application.
  async initialize() {
    // If auth is not working, enable debug logging to help diagnose.
    this.pca = await createNestablePublicClientApplication(getMsalConfig(true));
  }

  /**
   *
   * Uses MSAL and nested app authentication to get an access token from Office SSO.
   * This demonstrates how to work with user identity from the token.
   *
   * @param scopes The minimum scopes needed.
   * @returns The access token.
   */
  async ssoGetAccessToken(scopes: string[]): Promise<string | undefined> {
    // Check if access token is already stored.
    if (this.gUserProfile.accessToken) {
      return this.gUserProfile.accessToken;
    } else {
      const userProfile = await this.ssoGetUserIdentity();
      return userProfile.accessToken;
    }
  }

  async getTokenWithDialogApi(isInternetExplorer?: boolean): Promise<UserProfile> {
    return new Promise((resolve) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`${isInternetExplorer ? "dialogie.html" : "dialog.html"}`),
        (result) => {
          result.value.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (arg: { message: string; origin: string | undefined }) => {
              this.gUserProfile = JSON.parse(arg.message);
              resolve(this.gUserProfile);
              result.value.close();
            }
          );
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
   * @returns The user account data (including identity). The access token as string if falling back to dialog API.
   */
  async ssoGetUserIdentity(): Promise<UserProfile> {
    let userAccount: AuthenticationResult | undefined;

    // Return global user profile if already stored.
    if (this.gUserProfile.accessToken) {
      return this.gUserProfile;
    }

    if (!this.pca) {
      throw new Error("AccountManager is not initialized!");
    }
    const tokenRequest = {
      scopes: ["user.read"],
    };

    try {
      console.log("Trying to acquire token silently...");
      const authResult = await this.pca.acquireTokenSilent(tokenRequest);
      console.log("Acquired token silently.");
      const idTokenClaims = authResult.idTokenClaims as { name?: string; preferred_username?: string };
      this.gUserProfile = {
        userName: idTokenClaims.name,
        userEmail: idTokenClaims.preferred_username,
        accessToken: authResult.accessToken,
      };
      userAccount = authResult;
    } catch (error) {
      console.log(`Unable to acquire token silently: ${error}`);
    }

    if (userAccount === undefined) {
      // Acquire token silent failure. Send an interactive request via popup.
      try {
        if (this.fallbackPopup) {
          // Fall back to popup workaround for about:blank popup bug.
          this.gUserProfile = await this.getTokenWithDialogApi();
        } else {
          console.log("Trying to acquire token interactively...");
          const authResult = await this.pca.acquireTokenPopup(tokenRequest);
          console.log("Acquired token interactively.");
          userAccount = authResult;
        }
      } catch (popupError) {
        // Optional fallback if about:blank popup should not be shown
        if (popupError instanceof BrowserAuthError && popupError.errorCode === "popup_window_error") {
          this.fallbackPopup = true;
          this.gUserProfile = await this.getTokenWithDialogApi();
        } else {
          // Acquire token interactive failure.
          console.log(`Unable to acquire token interactively: ${popupError}`);
          throw new Error(`Unable to acquire access token: ${popupError}`);
        }
      }
    }
    return this.gUserProfile;
  }
}
