/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, fetch */

import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

// TODO: Replace with your Entra ID app registration's Application (client) ID.
const CLIENT_ID = "YOUR-ENTRA-ID-APP-REGISTRATION-CLIENT-ID";

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/common",
  },
  cache: {
    cacheLocation: "localStorage",
  },
};

const tokenRequest = { scopes: ["Mail.Send"] };

let msalInstance;
let userEmail = "";

Office.onReady(async () => {
  userEmail = Office.context.mailbox.userProfile.emailAddress;

  document.getElementById("recipient").textContent = userEmail;
  document.getElementById("context-input").value = JSON.stringify(generateSampleContext(), null, 2);
  document.getElementById("send-btn").addEventListener("click", handleSend);

  try {
    msalInstance = await createNestablePublicClientApplication(msalConfig);
  } catch (e) {
    document.getElementById("status").textContent = "Auth init failed: " + e.message;
    document.getElementById("status").className = "error";
  }
});

function generateSampleContext() {
  return {
    dateStamp: new Date().toISOString(),
    myStringProperty: "Hello world",
    myBooleanProperty: true,
  };
}

function generateActionableMessageBody(context) {
  const facts = Object.entries(context).map(([name, value]) => ({ title: name, value: String(value) }));

  const adaptiveCard = {
    type: "AdaptiveCard",
    version: "1.0",
    originator: "527104a1-f1a5-475a-9199-7a968161c870",
    hideOriginalBody: true,
    body: [
      {
        type: "TextBlock",
        text: "Activate **Actionable Message Activation** add-in",
        size: "medium",
        weight: "bolder",
      },
      {
        type: "FactSet",
        facts,
      },
      {
        type: "ActionSet",
        actions: [
          {
            type: "Action.InvokeAddInCommand",
            title: "View Initialization Context",
            addInId: "527104a1-f1a5-475a-9199-7a968161c870",
            desktopCommandId: "showInitContext",
            initializationContext: context,
          },
        ],
      },
    ],
  };

  return (
    "<html><head>" +
    '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">' +
    '<script type="application/ld+json">' +
    JSON.stringify(adaptiveCard) +
    "</script></head><body>" +
    "<p>If you don't see a message card above with clickable buttons, " +
    "your email client doesn't support Actionable Messages. Please try viewing " +
    "this mail in Outlook on the web for Office 365, or the latest version of Outlook 2016 for Windows.</p>" +
    "</body></html>"
  );
}

async function handleSend() {
  const statusEl = document.getElementById("status");
  const sendBtn = document.getElementById("send-btn");

  if (!msalInstance) {
    statusEl.textContent = "Authentication is not available. NAA may not be supported in this environment.";
    statusEl.className = "error";
    return;
  }

  // Parse context JSON.
  let context;
  try {
    context = JSON.parse(document.getElementById("context-input").value);
  } catch (e) {
    statusEl.textContent = "Invalid JSON: " + e.message;
    statusEl.className = "error";
    return;
  }

  sendBtn.disabled = true;
  statusEl.textContent = "Authenticating...";
  statusEl.className = "";

  let accessToken;
  try {
    const response = await msalInstance.acquireTokenSilent(tokenRequest);
    accessToken = response.accessToken;
  } catch (silentError) {
    if (silentError instanceof InteractionRequiredAuthError) {
      try {
        const response = await msalInstance.acquireTokenPopup(tokenRequest);
        accessToken = response.accessToken;
      } catch (popupError) {
        statusEl.textContent = "Authentication failed: " + popupError.message;
        statusEl.className = "error";
        sendBtn.disabled = false;
        return;
      }
    } else {
      statusEl.textContent = "Authentication failed: " + silentError.message;
      statusEl.className = "error";
      sendBtn.disabled = false;
      return;
    }
  }

  await sendEmail(accessToken, context);
}

async function sendEmail(accessToken, context) {
  const statusEl = document.getElementById("status");
  const sendBtn = document.getElementById("send-btn");

  sendBtn.disabled = true;
  statusEl.textContent = "Sending email...";
  statusEl.className = "";

  const htmlBody = generateActionableMessageBody(context);

  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendmail", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: {
          subject: "ACTIVATE YOUR ADD-IN",
          body: {
            contentType: "HTML",
            content: htmlBody,
          },
          toRecipients: [{ emailAddress: { address: userEmail } }],
        },
      }),
    });

    if (response.ok || response.status === 202) {
      statusEl.textContent = "Actionable message sent successfully!";
      statusEl.className = "success";
      sendBtn.disabled = false;
    } else {
      const error = await response.json();
      statusEl.textContent = "Send failed: " + (error.error?.message || response.statusText);
      statusEl.className = "error";
      sendBtn.disabled = false;
    }
  } catch (e) {
    statusEl.textContent = "Send failed: " + e.message;
    statusEl.className = "error";
    sendBtn.disabled = false;
  }
}
