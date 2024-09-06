/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, fetch, Office */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { auth } from "./authconfig";

let pca = undefined;
let isPCAInitialized = false;

// Called when loaded into Outlook on web.
Office.onReady(() => {});

/**
 * Initialize the public client application to work with SSO through NAA.
 */
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
    // All console.log statements write to the runtime log. For more information, see https://learn.microsoft.com/office/dev/add-ins/testing/runtime-logging
    console.log(`Error creating pca: ${error}`);
  }
}

/**
 * Gets the user name from Microsoft Graph. Uses an access token acquired through NAA and SSO.
 * @returns the user name (display name).
 */
async function getUserName() {
  await initializePCA();
  // Specify minimum scopes needed for the access token.
  const tokenRequest = {
    scopes: ["User.Read", "openid", "profile"],
  };
  let accessToken = null;

  // Acquire the access token silently.
  try {
    console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    console.log(`Unable to acquire token silently: ${error}`);
    throw error;
  }

  // Throw an error if the token is still null.
  if (accessToken === null) {
    console.log(`Unable to acquire access token. Access token is null.`);
    throw new Error("Unable to acquire access token. Access token is null.");
  }

  // Call the Microsoft Graph API with the access token.
  const response = await fetch("https://graph.microsoft.com/v1.0/me", {
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

/**
 * Called when the user creates a new email. Will set the signature using the signed-in user's name.
 * @param {*} event The event context from Office.
 */
function onNewMessageComposeHandler(event) {
  setSignature(event);
}

/**
 * Called when the user creates a new appointment. Will set the signature using the signed-in user's name.
 * @param {*} event The event context from Office.
 */
function onNewAppointmentComposeHandler(event) {
  setSignature(event);
}

/**
 * Sets the signature in the email item to indicate it is from the signed-in user.
 * @param {*} event The event context from Office.
 */
async function setSignature(event) {
  const item = Office.context.mailbox.item;

  // Add the signature.
  try {
    const userName = await getUserName();
    item.body.setSignatureAsync(
      "<b>From the desk of " + userName + ".",
      { asyncContext: event, coercionType: Office.CoercionType.Html },
      addSignatureCallback
    );
  } catch (error) {
    notifyUserToSignIn();
  }
  event.completed();
}

/**
 * Callback function to handle the result of adding a signature to the mail item.
 * @param {*} result The result from attemting to set the signature
 */
function addSignatureCallback(result) {
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.log(result.error.message);
  } else {
    console.log("Successfully added signature.");
    result.asyncContext.completed();
  }
}

/**
 * Gets correct command id to match to item type (appointment or message).
 * @returns The command id.
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Adds a notification to the email item requesting the user to sign in using the task pane.
 */
function notifyUserToSignIn() {
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

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
