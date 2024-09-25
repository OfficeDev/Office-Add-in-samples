// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides generic auth types and methods for task pane code (taskpane.ts).
// It handles any compatibility switching based on Edge vs Trident IE11 webviews.

/* global Office, window, document */

import { createLocalUrl } from "./util";
import { type AccountManager } from "./authConfig";

/**
 * Represents a user profile from an MSAL account.
 */
export interface UserProfile {
  userName?: string;
  userEmail?: string;
  accessToken?: string;
}

export enum AuthMethod {
  NAA, // Edge webview (Nested App Auth).
  MSALV2, // For Trident IE11 webview support.
}

// Globals
let msalAccountManager: AccountManager | undefined; // Account manager for MSAL v3 (NAA). Imported dynamically only if needed.
let authMethod = AuthMethod.NAA; // Default to nested app authentication.
let userProfile: UserProfile = {};

/**
 * Initializes which authentication method to use based on which webview is in use (Edge or Trident IE11.)
 */
export async function initializeAuthMethod() {
  // Check if Trident IE11 webview is in use.
  if (window.navigator.userAgent.indexOf("Trident") !== -1) {
    // Fall back to MSAL v2 to support Trident IE11 webview.
    authMethod = AuthMethod.MSALV2;
  } else {
    // Initialize the MSAL v3 (NAA) library.
    // Dynamically import the auth config code (MSAL v3 won't load in Trident IE11 webview.)
    const accountModule = await import("./authConfig");
    msalAccountManager = new accountModule.AccountManager();
    msalAccountManager.initialize();
  }
}

/**
 * Gets an access token for requested scopes. Handles switching if fallback auth is in use.
 */
export async function getAccessToken(scopes: string[]) {
  switch (authMethod) {
    case AuthMethod.NAA:
      // Use the MSAL v3 NAA library.
      if (msalAccountManager === undefined) throw new Error("msalAccountManager was not initialized!");
      userProfile.accessToken = await msalAccountManager.ssoGetAccessToken(scopes);
      break;
    case AuthMethod.MSALV2:
      // If Trident IE11 webview is in use, call getUserProfileWithDialogApi() to use the MSAL v2 compatible library.
      userProfile = await getUserProfileWithDialogApi();
      break;
  }
  if (userProfile.accessToken) {
    return userProfile.accessToken;
  } else {
    throw new Error("Could not get access token!");
  }
}

/**
 * Gets user profile information (name, email, access token). Handles switching if fallback auth is in use.
 */
export async function getUserProfile(): Promise<UserProfile> {
  switch (authMethod) {
    case AuthMethod.NAA:
      // Use the MSAL v3 NAA library.
      userProfile = await msalAccountManager?.ssoGetUserIdentity();
      break;
    case AuthMethod.MSALV2:
      // IE11 webview is in use. Call getUserProfileWithDialogApi to use the MSAL v2 compatible library.
      userProfile = await getUserProfileWithDialogApi();
      setSignOutButtonVisibility(true);
      break;
  }
  return userProfile;
}

/**
 * Uses the Office Dialog API to open an MSAL v2 auth window to sign in the user.
 * Used for Trident IE11 webview compatibility.
 * @returns The access token for the signed in user.
 */
export async function getUserProfileWithDialogApi(): Promise<UserProfile> {
  //TODO add check here to be sure not called twice in a row.
  //lodash has a way to do this.
  // store the promise in a state variable, and if populated just return the promise

  // Return token if already stored.
  // Note: does not handle case where token expires.
  if (userProfile.accessToken) {
    return userProfile;
  }

  // Encapsulate the dialog API call in a Promise.
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(createLocalUrl("dialogie.html"), { height: 60, width: 30 }, (result) => {
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
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      createLocalUrl("signoutdialogie.html"),
      { height: 60, width: 30 },
      (result) => {
        result.value.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: { message: string; origin: string | undefined }) => {
            userProfile = {};
            setSignOutButtonVisibility(false);
            result.value.close();
            resolve();
          }
        );
      }
    );
  });
}

/**
 * Makes the Sign out button visible or invisible on the task pane.
 *
 * @param visible true if the sign out button should be visible; otherwise, false.
 * @returns
 */
export function setSignOutButtonVisibility(visible: boolean) {
  const signOutButton = document.getElementById("signOutButton");
  if (!signOutButton) return;
  if (visible) {
    signOutButton.classList.remove("is-disabled");
  } else {
    signOutButton.classList.add("is-disabled");
  }
}
