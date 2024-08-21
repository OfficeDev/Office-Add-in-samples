/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { createNestablePublicClientApplication } from "@azure/msal-browser";

const sideloadMsg = document.getElementById("sideload-msg");
const appBody = document.getElementById("app-body");
const signInButton = document.getElementById("btnSignIn");
const itemSubject = document.getElementById("item-subject");

let pca = undefined;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    if (signInButton) {
      signInButton.onclick = signInUser;
    }
    // Initialize the public client application
    pca = await createNestablePublicClientApplication({
      auth: {
        clientId: "Enter_the_Application_Id_Here",
        authority: "https://login.microsoftonline.com/common"
      },
    });
  }
});

async function signInUser() {
  // Specify minimum scopes needed for the access token.
  const tokenRequest = {
    scopes: ["User.Read", "openid", "profile"],
  };
  let accessToken = null;

  // Call acquireTokenSilent.
  try {
    console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    console.log(`Unable to acquire token silently: ${error}`);
  }
  // Call acquireTokenPopup.
  if (accessToken === null) {
    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const userAccount = await pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      accessToken = userAccount.accessToken;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Unable to acquire token interactively: ${popupError}`);
    }
  }
  // Log error if both silent and popup requests failed.
  if (accessToken === null) {
    console.error(`Unable to acquire access token.`);
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
