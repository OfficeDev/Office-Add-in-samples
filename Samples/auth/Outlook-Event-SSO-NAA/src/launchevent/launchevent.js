/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { auth } from "./authconfig";

let pca = undefined;
let isPCAInitialized = false;

async function initializePCA() {
  if (isPCAInitialized) {
    return;
  }

  // Initialize the public client application.
  try {
    pca = await createNestablePublicClientApplication({
      auth: auth,
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
    throw error; //rethrow
  }

  // Log an error if the token is still null.
  if (accessToken === null) {
    console.log(`Unable to acquire access token. Access token is null.`);
    throw new Error("Unable to acquire access token. Access token is null.");
    return;
  }

  // Call the Microsoft Graph API with the access token.
  const response = await fetch(`https://graph.microsoft.com/v1.0/me`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    // Get the username from the response JSON.
    const data = await response.json();
    const name = data.displayName;

    return name;
  } else {
    const errorText = await response.text();
    console.log("Microsoft Graph call failed - error text: " + errorText);
  }
}

function onNewMessageComposeHandler(event) {
  //addInsight();
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
      try {
        const userName = await getUserName();
        item.body.setSignatureAsync(
          "<b>From the desk of " + userName + ".",
          { asyncContext: result.asyncContext, coercionType: Office.CoercionType.Html },
          addSignatureCallback
        );
      } catch (error) {
        addInsight();
      }
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

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

function addInsight() {
  Office.context.mailbox.item.notificationMessages.addAsync("16c028c6_sign_in_notification", {
    type: "insightMessage",
    message: "Please sign in using the task pane to start using the Office Add-ins sample.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Sign in",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform === null) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
}
