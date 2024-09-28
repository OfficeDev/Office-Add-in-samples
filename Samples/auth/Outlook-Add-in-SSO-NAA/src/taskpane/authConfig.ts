// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global Office, console, document*/

import {
  BrowserAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { msalConfig } from "./msalconfig";
import { createLocalUrl } from "./util";
import { getTokenRequest } from "./msalcommon";

export type AuthDialogResult = {
  accessToken?: string;
  error?: string;
};

type DialogEventMessage = { message: string; origin: string | undefined };
type DialogEventError = { error: number };
type DialogEventArg = DialogEventMessage | DialogEventError;

// Encapsulate functions for getting user account and token information.
export class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private _dialogApiResult: Promise<string> | null = null;
  private _usingFallbackDialog = false;

  private getSignOutButton() {
    return document.getElementById("signOutButton");
  }

  private setSignOutButtonVisibility(isVisible: boolean) {
    const signOutButton = this.getSignOutButton();
    if (signOutButton) {
      signOutButton.style.visibility = isVisible ? "visible" : "hidden";
    }
  }

  private isNestedAppAuthSupported() {
    return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  }

  // Initialize MSAL public client application.
  async initialize() {
    // Make sure office.js is initialized
    await Office.onReady();

    // If auth is not working, enable debug logging to help diagnose.
    this.pca = await createNestablePublicClientApplication(msalConfig);

    // If Office does not support Nested App Auth provide a sign-out button since the user selects account
    if (!this.isNestedAppAuthSupported() && this.pca.getActiveAccount()) {
      this.setSignOutButtonVisibility(true);
    }
    this.getSignOutButton()?.addEventListener("click", () => this.signOut());
  }

  private async signOut() {
    if (this._usingFallbackDialog) {
      await this.signOutWithDialogApi();
    } else {
      await this.pca?.logoutPopup();
    }

    this.setSignOutButtonVisibility(false);
  }

  /**
   *
   * @param scopes the minimum scopes needed.
   * @returns An access token.
   */
  async ssoGetAccessToken(scopes: string[]) {
    if (this._dialogApiResult) {
      return this._dialogApiResult;
    }

    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }

    try {
      console.log("Trying to acquire token silently...");
      const authResult = await this.pca.acquireTokenSilent(getTokenRequest(scopes, false));
      console.log("Acquired token silently.");
      return authResult.accessToken;
    } catch (error) {
      console.warn(`Unable to acquire token silently: ${error}`);
    }

    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const selectAccount = this.pca.getActiveAccount() ? false : true;
      const authResult = await this.pca.acquireTokenPopup(getTokenRequest(scopes, selectAccount));
      console.log("Acquired token interactively.");
      if (selectAccount) {
        this.pca.setActiveAccount(authResult.account);
      }
      if (!this.isNestedAppAuthSupported()) {
        this.setSignOutButtonVisibility(true);
      }
      return authResult.accessToken;
    } catch (popupError) {
      // Optional fallback if about:blank popup should not be shown
      if (popupError instanceof BrowserAuthError && popupError.errorCode === "popup_window_error") {
        const accessToken = await this.getTokenWithDialogApi();
        return accessToken;
      } else {
        // Acquire token interactive failure.
        console.error(`Unable to acquire token interactively: ${popupError}`);
        throw new Error(`Unable to acquire access token: ${popupError}`);
      }
    }
  }

  /**
   * Gets an access token by using the Office dialog API to handle authentication. Used for fallback scenario.
   * @returns The access token.
   */
  async getTokenWithDialogApi(): Promise<string> {
    this._dialogApiResult = new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(createLocalUrl(`dialog.html`), { height: 60, width: 30 }, (result) => {
        result.value.addEventHandler(Office.EventType.DialogEventReceived, (arg: DialogEventArg) => {
          const errorArg = arg as DialogEventError;
          if (errorArg.error == 12006) {
            this._dialogApiResult = null;
            reject("Dialog closed");
          }
        });
        result.value.addEventHandler(Office.EventType.DialogMessageReceived, (arg: DialogEventArg) => {
          const messageArg = arg as DialogEventMessage;
          const parsedMessage = JSON.parse(messageArg.message);
          result.value.close();

          if (parsedMessage.error) {
            reject(parsedMessage.error);
            this._dialogApiResult = null;
          } else {
            resolve(parsedMessage.accessToken);
            this.setSignOutButtonVisibility(true);
            this._usingFallbackDialog = true;
          }
        });
      });
    });
    return this._dialogApiResult;
  }

  signOutWithDialogApi(): Promise<void> {
    return new Promise((resolve) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html?logout=1`),
        { height: 60, width: 30 },
        (result) => {
          result.value.addEventHandler(Office.EventType.DialogMessageReceived, () => {
            this.setSignOutButtonVisibility(false);
            this._dialogApiResult = null;
            resolve();
            result.value.close();
          });
        }
      );
    });
  }
}
