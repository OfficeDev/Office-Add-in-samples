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
import { UserProfile, setSignOutButtonVisibility } from "./authhelper";

export { AccountManager };

// Encapsulate functions for getting user account and token information.
class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private fallbackPopup = false; // true if Outlook client has the about:blank popup bug and we need to fall back.
  private userProfile: UserProfile = {}; // Tracks the user profile including access token.

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
    if (this.userProfile.accessToken) {
      return this.userProfile.accessToken;
    } else {
      const userProfile = await this.ssoGetUserIdentity(scopes);
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
              this.userProfile = JSON.parse(arg.message);
              resolve(this.userProfile);
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
  async ssoGetUserIdentity(scopes: string[]): Promise<UserProfile> {
    // Return global user profile if already stored.
    if (this.userProfile.accessToken) {
      return this.userProfile;
    }

    if (!this.pca) {
      throw new Error("AccountManager is not initialized!");
    }
    const tokenRequest = {
      scopes,
    };

    let userAccount: AuthenticationResult | undefined;

    try {
      console.log("Trying to acquire token silently...");
      const authResult = await this.pca.acquireTokenSilent(tokenRequest);
      console.log("Acquired token silently.");
      const idTokenClaims = authResult.idTokenClaims as { name?: string; preferred_username?: string };
      this.userProfile = {
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
          this.userProfile = await this.getTokenWithDialogApi();
          setSignOutButtonVisibility(true);
        } else {
          console.log("Trying to acquire token interactively...");
          userAccount = await this.pca.acquireTokenPopup(tokenRequest);
          console.log("Acquired token interactively.");
        }
      } catch (popupError) {
        // Optional fallback if about:blank popup should not be shown
        if (popupError instanceof BrowserAuthError && popupError.errorCode === "popup_window_error") {
          this.fallbackPopup = true;
          this.userProfile = await this.getTokenWithDialogApi();
        } else {
          // Acquire token interactive failure.
          console.log(`Unable to acquire token interactively: ${popupError}`);
          throw new Error(`Unable to acquire access token: ${popupError}`);
        }
      }
    }

    // Create user profile from MSAL user account.
    if (userAccount !== undefined) {
      const idTokenClaims = userAccount.idTokenClaims as { name?: string; preferred_username?: string };
      this.userProfile = {
        userName: idTokenClaims.name,
        userEmail: idTokenClaims.preferred_username,
        accessToken: userAccount.accessToken,
      };
      return this.userProfile;
    }
  }
}
