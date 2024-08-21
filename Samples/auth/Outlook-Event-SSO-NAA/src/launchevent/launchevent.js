/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { createNestablePublicClientApplication } from "@azure/msal-browser";

let pca = undefined;
let isPCAInitialized = false;

async function initializePCA() {
  if (isPCAInitialized) return;

  // Initialize the public client application
  try {
    pca = await createNestablePublicClientApplication({
      auth: {
        clientId: "605f8396-522e-4d3c-a83d-829fd2fcf47e", //Enter_the_Application_Id_Here
        authority: "https://login.microsoftonline.com/common",
      },
    });
    isPCAInitialized = true;
  } catch (error) {
    console.log(`Error creating pca: ${error}`);
  }
}

async function getUserName() {
  await initializePCA();
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

  //Log error if token still null.
  if (accessToken === null) {
    console.log(`Unable to acquire access token. Access token is null.`);
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

    return name;
  } else {
    const errorText = await response.text();
    console.log("Microsoft Graph call failed - error text: " + errorText);
  }
}

function onNewMessageComposeHandler(event) {
  setSignature(event);
}
function onNewAppointmentComposeHandler(event) {
  setSignature(event);
}
async function setSignature(event) {
  const item = Office.context.mailbox.item;

  // Check if a default Outlook signature is already configured.
  item.isClientSignatureEnabledAsync({ asyncContext: event }, async (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log(result.error.message);
      return;
    }

    // Add a signature if there's no default Outlook signature configured.
    if (result.value === false) {
      const userName = await getUserName();
      item.body.setSignatureAsync(
        "<b>From the desk of " + userName + ".",
        { asyncContext: result.asyncContext, coercionType: Office.CoercionType.Html },
        addSignatureCallback
      );
    }
  });
}

// Callback function to add a signature to the mail item.
function addSignatureCallback(result) {
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.log(result.error.message);
    return;
  }

  console.log("Successfully added signature.");
  result.asyncContext.completed();
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
}
