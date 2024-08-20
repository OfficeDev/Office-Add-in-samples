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
        clientId: "[Enter-app-registration-id-here]",
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
    scopes: ["Files.Read", "User.Read", "openid", "profile"],
  };
  let accessToken = null;

  // TODO 1: Call acquireTokenSilent.
  try {
    console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    console.log(`Unable to acquire token silently: ${error}`);
  }
  // TODO 2: Call acquireTokenPopup.
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
  // TODO 3: Log error if token still null.
  // Log error if both silent and popup requests failed.
  if (accessToken === null) {
    // console.error(`Unable to acquire access token.`);
    return;
  }
  // TODO 4: Call the Microsoft Graph API.
  // Call the Microsoft Graph API with the access token.
  const response = await fetch(`https://graph.microsoft.com/v1.0/me`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    // Write file names to the console.
    const data = await response.json();
    console.log("data is " + JSON.stringify(data));
    const name = data.displayName;

    // Be sure the taskpane.html has an element with Id = item-subject.
    // const label = document.getElementById("item-subject");

    // Write file names to task pane and the console.
    //const nameText = names.join(", ");
    //if (label) label.textContent = nameText;
    // console.log(nameText);
    return name;
  } else {
    const errorText = await response.text();
    console.error("Microsoft Graph call failed - error text: " + errorText);
  }
}

function onNewMessageComposeHandler(event) {
  console.log("OnNewMessageComposeHandler");
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

/**
 * Creates information bar to display when new message or appointment is created
 */
// Not used at this time.
function display_insight_infobar() {
  console.log("display insight infobar");
  Office.context.mailbox.item.notificationMessages.addAsync("16c028c6-f97d-4b09-96eb-3821219e0a47", {
    type: "insightMessage",
    message: "Add-in unable to process events. Please sign in using the task pane.",
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

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  console.log("getting command id");
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}


// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
}
