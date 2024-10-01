// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides generic auth types and methods for task pane code (taskpane.ts).
// It handles any compatibility switching based on Edge vs Trident IE11 webviews.

/* global Office, document */

import { createLocalUrl, isInternetExplorer } from "./util";
import { type AccountManager } from "./msalAuth";
import { FallbackError } from "./errors";

type DialogEventMessage = { message: string; origin: string | undefined };
type DialogEventError = { error: number };
type DialogEventArg = DialogEventMessage | DialogEventError;

export type AuthDialogResult = {
  accessToken?: string;
  error?: string;
};

export enum AuthMethod {
  NAA, // Edge webview (Nested App Auth).
  MSALV2DialogApi, // For Trident IE11 webview support.
}

// Globals
let msalAccountManager: AccountManager | undefined; // Account manager for MSAL v3 (NAA). Imported dynamically only if needed.
let authMethod = AuthMethod.NAA; // Default to nested app authentication.
let dialogApiResult: Promise<string> | null = null;

/**
 * Initializes which authentication method to use based on which webview is in use (Edge or Trident IE11.)
 */
export async function initializeAuthMethod() {
  // Check if Trident IE11 webview is in use.
  if (isInternetExplorer()) {
    // Fall back to MSAL v2 to support Trident IE11 webview.
    authMethod = AuthMethod.MSALV2DialogApi;
  } else {
    // Initialize the MSAL v3 (NAA) library.
    // Dynamically import the auth config code (MSAL v3 won't load in Trident IE11 webview.)
    const accountModule = await import("./msalAuth");
    msalAccountManager = new accountModule.AccountManager();
    await msalAccountManager.initialize();
    if (msalAccountManager.hasActiveAccount() && !msalAccountManager.isNestedAppAuthSupported()) {
      setSignOutButtonVisibility(true);
    }
  }
  getSignOutButton()?.addEventListener("click", () => signOut());
}

function getSignInDialogUrl() {
  if (authMethod === AuthMethod.MSALV2DialogApi) {
    return createLocalUrl("dialoginternetexplorer.html");
  } else {
    return createLocalUrl("dialog.html");
  }
}

function getSignOutDialogUrl() {
  if (authMethod === AuthMethod.MSALV2DialogApi) {
    return createLocalUrl("signoutdialoginternetexplorer.html");
  } else {
    return createLocalUrl("signoutdialog.html");
  }
}

/**
 * Gets an access token for requested scopes. Handles switching if fallback auth is in use.
 */
export async function getAccessToken(scopes: string[]): Promise<string> {
  switch (authMethod) {
    case AuthMethod.NAA:
      return getTokenWithMsal(scopes);
    case AuthMethod.MSALV2DialogApi:
      // If Trident IE11 webview is in use, call getUserProfileWithDialogApi() to use the MSAL v2 compatible library.
      return getTokenWithDialogApi();
  }
}

async function getTokenWithMsal(scopes: string[]): Promise<string> {
  if (dialogApiResult) {
    return dialogApiResult;
  }
  if (msalAccountManager === undefined) throw new Error("msalAccountManager was not initialized!");

  let token = "";

  try {
    token = await msalAccountManager.ssoGetAccessToken(scopes);
    if (!msalAccountManager.isNestedAppAuthSupported()) {
      setSignOutButtonVisibility(true);
    }
  } catch (ex) {
    if (ex instanceof FallbackError) {
      token = await getTokenWithDialogApi();
    } else {
      throw ex;
    }
  }

  return token;
}
/**
 * Uses the Office Dialog API to open an MSAL v2 auth window to sign in the user.
 * Used for Trident IE11 webview compatibility.
 * @returns The access token for the signed in user.
 */
export async function getTokenWithDialogApi(): Promise<string> {
  if (dialogApiResult) {
    return dialogApiResult;
  }

  dialogApiResult = new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(getSignInDialogUrl(), { height: 60, width: 30 }, (result) => {
      result.value.addEventHandler(Office.EventType.DialogEventReceived, (arg: DialogEventArg) => {
        const errorArg = arg as DialogEventError;
        if (errorArg.error == 12006) {
          dialogApiResult = null;
          reject("Dialog closed");
        }
      });
      result.value.addEventHandler(Office.EventType.DialogMessageReceived, (arg: DialogEventArg) => {
        const messageArg = arg as DialogEventMessage;
        const parsedMessage = JSON.parse(messageArg.message);
        result.value.close();

        if (parsedMessage.error) {
          reject(parsedMessage.error);
          dialogApiResult = null;
        } else {
          resolve(parsedMessage.accessToken);
          setSignOutButtonVisibility(true);
        }
      });
    });
  });
  return dialogApiResult;
}

async function signOutWithDialogApi(): Promise<void> {
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(getSignOutDialogUrl(), { height: 60, width: 30 }, (result) => {
      result.value.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        resolve();
        result.value.close();
      });
    });
  });
}

async function signOut() {
  if (msalAccountManager && !dialogApiResult) {
    try {
      await msalAccountManager.signOut();
    } catch {
      await signOutWithDialogApi();
    }
  } else {
    dialogApiResult = null;
    await signOutWithDialogApi();
  }
  setSignOutButtonVisibility(false);
}

function getSignOutButton() {
  return document.getElementById("signOutButton");
}

/**
 * Makes the Sign out button visible or invisible on the task pane.
 *
 * @param isVisible true if the sign out button should be visible; otherwise, false.
 * @returns
 */
function setSignOutButtonVisibility(isVisible: boolean) {
  const signOutButton = getSignOutButton();
  if (signOutButton) {
    signOutButton.style.visibility = isVisible ? "visible" : "hidden";
  }
}
