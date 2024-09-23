/* global Office, window */

import { AuthMethod, UserProfile } from "./userProfile";
import { createLocalUrl } from "./util";

let msalAccountManager; // Account manager for MSAL v3 (NAA).
let authMethod = AuthMethod.NAA; // Default to nested app authentication.
let userProfile: UserProfile = {};

export async function initializeAuth() {
  // Check if Trident IE11 webview is in use.
  if (window.navigator.userAgent.indexOf("Trident") !== -1) {
    authMethod = AuthMethod.MSALV2;
  } else {
    // Initialize the MSAL v3 (NAA) library.
    // This will dynamically import the auth config code (MSAL v3 won't load in the IE11 webview.)
    const accountModule = await import("./authConfig");
    msalAccountManager = new accountModule.AccountManager();
    msalAccountManager.initialize();
  }
}

/**
 * Gets an access token for requested scopes. Handles switching if fallback auth is in use.
 */
export async function getAccessToken(scopes: string[]) {
  let userProfile: UserProfile = {};
  switch (authMethod) {
    case AuthMethod.NAA:
      // Use the MSAL v3 NAA library.
      userProfile.accessToken = msalAccountManager.ssoGetAccessToken(scopes);
      break;
    case AuthMethod.MSALV2:
      // If IE11 webview is in use, call getTokenWithDialogApi(true) to use the MSAL v2 compatible library.
      userProfile = await getTokenWithDialogApi(scopes);
      break;
  }
  if (userProfile.accessToken) {
    return userProfile.accessToken;
  } else {
    throw new Error("Could not get access token!");
  }
}

/**
 * Gets an access token for requested scopes. Handles switching if fallback auth is in use.
 */
export async function getUserProfile(): Promise<UserProfile> {
  switch (authMethod) {
    case AuthMethod.NAA:
      // Use the MSAL v3 NAA library.
      userProfile = msalAccountManager.ssoGetUserIdentity();
      break;
    case AuthMethod.MSALV3:
      // Use the MSAL v3 library but with the Office dialog API (non-sso fallback).
      userProfile = await getTokenWithDialogApi();
      break;
    case AuthMethod.MSALV2:
      // If IE11 webview is in use, call getTokenWithDIalogApi(true) to use the MSAL v2 compatible library.
      userProfile = await getTokenWithDialogApi();
      break;
  }
  return userProfile;
}

/**
 * Uses the Office Dialog API to open an MSAL auth window to sign in the user.
 * @param isInternetExplorer true if add-in hosted in the IE11 webview; otherwise, false.
 * @returns The access token for the signed in user.
 */
export async function getTokenWithDialogApi(): Promise<UserProfile> {
  // Return token if already stored.
  if (userProfile.accessToken) {
    return userProfile;
  }

  // Encapsulate the dialog API call in a Promise.
  return new Promise((resolve) => {
    // Determine if dialog for IE 11 should be used for Trident webview.
    let dialogPage = "dialog.html";
    if (authMethod === AuthMethod.MSALV2) {
      dialogPage = "dialogie.html";
    }
    Office.context.ui.displayDialogAsync(createLocalUrl(dialogPage), { height: 60, width: 30 }, (result) => {
      result.value.addEventHandler(
        Office.EventType.DialogMessageReceived,
        (arg: { message: string; origin: string | undefined }) => {
          userProfile = JSON.parse(arg.message);
          resolve(userProfile);
          result.value.close();
        }
      );
    });
  });
}
/**
 * Sign out the user from MSAL.
 */
export async function signOutUser(): Promise<void> {
  let dialogPage = "";
  if (authMethod === AuthMethod.MSALV2) {
    dialogPage = "signoutdialogie.html"; // Use IE
  } else {
    dialogPage = "signoutdialog.html";
  }

  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(createLocalUrl(dialogPage), { height: 60, width: 30 }, (result) => {
      result.value.addEventHandler(
        Office.EventType.DialogMessageReceived,
        (arg: { message: string; origin: string | undefined }) => {
          const parsedMessage = JSON.parse(arg.message);
          resolve();
          result.value.close();
        }
      );
    });
  });
}
