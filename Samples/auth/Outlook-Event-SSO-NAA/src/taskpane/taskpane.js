/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, console, fetch, Office */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { auth } from "../launchevent/authconfig";

const sideloadMsg = document.getElementById("sideload-msg");
const signInButton = document.getElementById("btnSignIn");
const itemSubject = document.getElementById("item-subject");

let pca = undefined;
let isPCAInitialized = false;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    sideloadMsg.style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    signInButton.onclick = signInUser;
    // Initialize the public client application.
    try {
      pca = await createNestablePublicClientApplication({
        auth: auth,
      });
      isPCAInitialized = true;
    } catch (error) {
      // All console.log statements write to the runtime log. For more information, see https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging
      console.log(`Error creating pca: ${error}`);
    }
  }
});

/**
 * Signs in the user using NAA and SSO auth flow. If successful, displays the user's name in the task pane.
 */
async function signInUser() {
  // Check that PCA initialized correctly in Office.onReady().
  if (!isPCAInitialized) {
    itemSubject.innerText = "Can't sign in because the PCA could not be initialized. See console logs for details.";
    return;
  }

  // Specify minimum scopes needed for the access token.
  const tokenRequest = {
    scopes: ["User.Read"],
  };
  let accessToken = null;
  try {
    const authResult = await pca.acquireTokenSilent(tokenRequest);
    accessToken = authResult.accessToken;
    console.log("Acquired token silently.");
  } catch (error) {
    console.log(`Unable to acquire token silently: ${error}`);
  }
  if (accessToken === null) {
    // If silently acquiring the token fails, send an interactive request via popup.
    try {
      const authResult = await pca.acquireTokenPopup(tokenRequest);
      accessToken = authResult.accessToken;
      console.log("Acquired token interactively.");
    } catch (popupError) {
      // Failed to acquire the token with the popup.
      console.log(`Unable to acquire token interactively: ${popupError}`);
    }
  }

  // Log error if both silent and popup requests failed.
  if (accessToken === null) {
    console.error("Unable to acquire access token.");
    return;
  }

  // Call the Microsoft Graph API with the access token.
  const response = await fetch(`https://graph.microsoft.com/v1.0/me`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    // Get the user name from response JSON.
    const data = await response.json();
    const name = data.displayName;

    if (itemSubject) {
      itemSubject.innerText = "You are now signed in as " + name + ".";
    }
  } else {
    const errorText = await response.text();
    console.log("Microsoft Graph call failed - error text: " + errorText);
  }
}
