// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console, document*/

/// <reference types="office-js" />

import {
  BrowserAuthError,
  InteractionRequiredAuthError,
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

// Constants
const DIALOG_DIMENSIONS = { height: 60, width: 30 } as const;
const DIALOG_CLOSED_ERROR_CODE = 12006;
const POPUP_WINDOW_ERROR_CODE = "popup_window_error";
const SIGN_OUT_BUTTON_ID = "signOutButton";
const NESTED_APP_AUTH_REQUIREMENT = { name: "NestedAppAuth", version: "1.1" } as const;

// Encapsulate functions for getting user account and token information.
export class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private _dialogApiResult: Promise<string> | null = null;
  private _usingFallbackDialog = false;

  private setSignOutButtonVisibility(isVisible: boolean): void {
    const signOutButton = document.getElementById(SIGN_OUT_BUTTON_ID);
    if (signOutButton) {
      signOutButton.style.visibility = isVisible ? "visible" : "hidden";
    }
  }

  private isNestedAppAuthSupported(): boolean {
    return Office.context.requirements.isSetSupported(
      NESTED_APP_AUTH_REQUIREMENT.name, 
      NESTED_APP_AUTH_REQUIREMENT.version
    );
  }

  // Initialize MSAL public client application.
  async initialize(): Promise<void> {
    try {
      // Make sure office.js is initialized.
      await Office.onReady();

      // Initialize a nested public client application.
      this.pca = await createNestablePublicClientApplication(msalConfig);

      // If Office does not support nested app auth provide a sign-out button since the user selects account.
      if (!this.isNestedAppAuthSupported() && this.pca.getActiveAccount()) {
        this.setSignOutButtonVisibility(true);
      }
      
      // Add event listener for click event on sign out button.
      const signOutButton = document.getElementById(SIGN_OUT_BUTTON_ID);
      if (signOutButton) {
        signOutButton.addEventListener("click", () => this.signOut());
      }
    } catch (error) {
      console.error("Failed to initialize AccountManager:", error);
      throw new Error(`Initialization failed: ${error}`);
    }
  }

  // Sign out the user.
  private async signOut() {
    await (this._usingFallbackDialog ? this.signOutWithDialogApi() : this.pca?.logoutPopup());
    this.setSignOutButtonVisibility(false);
  }

  // Get login hint for Word, Excel, or PowerPoint on the web from the auth context.
  private async getLoginHint(): Promise<string | undefined> {
    try {
      if (typeof Office !== "undefined" && Office.context) {
            const authContext = await Office.auth.getAuthContext();
            if (authContext?.userPrincipalName) return authContext.userPrincipalName;
        }
    } catch (error) {
      console.warn("Could not get login hint:", error);
    }
    return undefined;
  }

  async acquireToken(scopes: string[]): Promise<string> {
    // Check if the user is already signed in via fallback dialog API.
    if (this._dialogApiResult) {
      return this._dialogApiResult;
    }
    
    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }
    const loginHint = await this.getLoginHint();
    console.log(loginHint);
    
    try {
      console.log("Trying to acquire token silently...");
      const tokenRequest = getTokenRequest(scopes, false, undefined, loginHint);
      // If we have a login hint, use SSO silent flow which is required for Word, Excel, or PowerPoint on the web.
      const authResult = loginHint 
        ? await this.pca!.ssoSilent(tokenRequest)
        : await this.pca!.acquireTokenSilent(tokenRequest);
      console.log("Acquired token silently.");
      return authResult.accessToken;
    } catch (silentError) {
      if (silentError instanceof InteractionRequiredAuthError) {
        return this.acquireTokenInteractively(scopes, loginHint);
      } else {
        // For running on a localhost server, use the following line of code
        // to work around CORS errors with localhost.
        // Comment this code when deploying to production.
        return this.acquireTokenInteractively(scopes, loginHint);

        // For production uncomment the following code.
        // throw new Error(`Unable to acquire access token: ${silentError}`);
        
      }
    }
  }

  private async acquireTokenInteractively(scopes: string[], loginHint: string | undefined): Promise<string> {
    try {
      console.log("Trying to acquire token interactively...");
      
      const authResult = await this.pca!.acquireTokenPopup(
        getTokenRequest(scopes, false, undefined, loginHint)
      );
      console.log("Acquired token interactively.");
      
      // Show sign-out button if Office doesn't support Nested App Auth
      if (!this.isNestedAppAuthSupported()) {
        this.setSignOutButtonVisibility(true);
      }
      return authResult.accessToken;
    } catch (popupError) {
      return this.handleInteractiveTokenError(popupError);
    }
  }

  private async handleInteractiveTokenError(popupError: unknown): Promise<string> {
    // Optional fallback if about:blank popup should not be shown
    if (popupError instanceof BrowserAuthError && popupError.errorCode === POPUP_WINDOW_ERROR_CODE) {
      const accessToken = await this.getTokenWithDialogApi();
      this.setSignOutButtonVisibility(true);
      return accessToken;
    } else {
      // Acquire token interactive failure.
      console.error(`Unable to acquire token interactively: ${popupError}`);
      throw new Error(`Unable to acquire access token: ${popupError}`);
    }
  }

  /**
   * Gets an access token by using the Office dialog API to handle authentication. Used for fallback scenario.
   * @returns The access token.
   */
  async getTokenWithDialogApi(): Promise<string> {
    this._dialogApiResult = new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html`), 
        DIALOG_DIMENSIONS, 
        (result: any) => {
          result.value.addEventHandler(Office.EventType.DialogEventReceived, (arg: DialogEventArg) => {
            if ((arg as DialogEventError).error === DIALOG_CLOSED_ERROR_CODE) {
              this._dialogApiResult = null;
              reject("Dialog closed");
            }
          });
          result.value.addEventHandler(Office.EventType.DialogMessageReceived, (arg: DialogEventArg) => {
            const parsedMessage = JSON.parse((arg as DialogEventMessage).message);
            result.value.close();
            if (parsedMessage.error) {
              this._dialogApiResult = null;
              reject(parsedMessage.error);
            } else {
              this.setSignOutButtonVisibility(true);
              this._usingFallbackDialog = true;
              resolve(parsedMessage.accessToken);
            }
          });
        }
      );
    });
    return this._dialogApiResult;
  }

  signOutWithDialogApi(): Promise<void> {
    return new Promise((resolve) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html?logout=1`), 
        DIALOG_DIMENSIONS, 
        (result: any) => {
          result.value.addEventHandler(Office.EventType.DialogMessageReceived, () => {
            this.setSignOutButtonVisibility(false);
            this._dialogApiResult = null;
            result.value.close();
            resolve();
          });
        }
      );
    });
  }

  /**
   * Clean up resources and event listeners
   */
  cleanup(): void {
    const signOutButton = document.getElementById(SIGN_OUT_BUTTON_ID);
    if (signOutButton) {
      signOutButton.removeEventListener("click", () => this.signOut());
    }
    this._dialogApiResult = null;
    this._usingFallbackDialog = false;
  }
}
