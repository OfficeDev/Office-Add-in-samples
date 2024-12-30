/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { createNestablePublicClientApplication, LogLevel } from "@azure/msal-browser";

/* global console */

let pca = undefined;

async function getAccessToken() {
  let accessToken = null;
  // Initialize the public client application
  if (!pca) {
    pca = await createNestablePublicClientApplication({
      auth: {
        //clientId: "c2d98363-620e-45ba-afff-b5fdb25a704e",
        clientId: "c3f801be-e126-49b2-a870-30c2940a4236",
        authority: "https://login.microsoftonline.com/common",
      },
      system: {
        loggerOptions: {
          logLevel: LogLevel.Verbose,
          loggerCallback: (level, message, containsPii) => {
            switch (level) {
              case LogLevel.Error:
                console.error(message);
                return;
              case LogLevel.Info:
                console.info(message);
                return;
              case LogLevel.Verbose:
                console.debug(message);
                return;
              case LogLevel.Warning:
                console.warn(message);
                return;
            }
          },
        },
      },
    });
  }
  // Specify minimum scopes needed for the access token.
  const tokenRequest = {
    //scopes: ["api://34ddb552-7075-4a62-b91f-84cbde89eab2/Contoso.Write"],
    scopes: ["api://d0ff5b49-5a37-4b63-889e-86cf8821eb07/Todolist.Read"],
    
    //scopes: ["user.read"],
  };

  try {
    console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    console.log(`Unable to acquire token silently: ${error}`);
  }

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
    throw new Error(`Unable to acquire access token.`);
  }
  return accessToken;
}

export async function setUserData(): Promise<void> {
  // get token from MSAL
  const accessToken = await getAccessToken();
  await callSetUserData(accessToken);
}

async function callSetUserData(accessToken) {
  // Call the Microsoft Graph API with the access token.
  const response = await fetch(`https://localhost:3000/setUserData`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    console.log("success");
  } else {
    const errorText = await response.text();
    console.error("failed - error text: " + errorText);
  }
  return response;
}
