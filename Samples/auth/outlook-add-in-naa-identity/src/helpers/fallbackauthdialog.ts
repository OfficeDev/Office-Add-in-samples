/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to your server and pass it to the task pane.
 */

/* global console, localStorage, location, Office, window */

import { Configuration, LogLevel, PublicClientApplication, RedirectRequest } from "@azure/msal-browser";
import { callGetUserData } from "./middle-tier-calls";
import { showMessage } from "./message-helper";

const clientId = "{application GUID here}"; //This is your client ID
const accessScope = `api://${window.location.host}/${clientId}/access_as_user`;
const loginRequest: RedirectRequest = {
  scopes: [accessScope],
  extraScopesToConsent: ["user.read"],
};

const msalConfig: Configuration = {
  auth: {
    clientId: clientId,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:{PORT}/fallbackauthdialog.html", // Update config script to enable `https://${window.location.host}/fallbackauthdialog.html`,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage", // Needed to avoid "User login is required" error.
    storeAuthStateInCookie: true, // Recommended to avoid certain IE/Edge issues.
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
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
};

const publicClientApp: PublicClientApplication = new PublicClientApplication(msalConfig);

let loginDialog: Office.Dialog = null;
let homeAccountId = null;
let callbackFunction = null;

Office.onReady(() => {
  if (Office.context.ui.messageParent) {
    publicClientApp
      .handleRedirectPromise()
      .then(handleResponse)
      .catch((error) => {
        console.log(error);
        Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
      });

    // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
    // stored login data in localStorage. So a direct call of acquireTokenRedirect
    // causes the error "User login is required". Once the user is logged in successfully
    // the first time, msal data in localStorage will prevent this error from ever hap-
    // pening again; but the error must be blocked here, so that the user can login
    // successfully the first time. To do that, call loginRedirect first instead of
    // acquireTokenRedirect.
    if (localStorage.getItem("loggedIn") === "yes") {
      publicClientApp.acquireTokenRedirect(loginRequest);
    } else {
      // This will login the user and then the (response.tokenType === "id_token")
      // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
      // and then the dialog is redirected back to this script, so the
      // acquireTokenRedirect above runs.
      publicClientApp.loginRedirect(loginRequest);
    }
  }
});

function handleResponse(response) {
  if (response.tokenType === "id_token") {
    console.log("LoggedIn");
    localStorage.setItem("loggedIn", "yes");
  } else {
    console.log("token type is:" + response.tokenType);
    Office.context.ui.messageParent(
      JSON.stringify({ status: "success", result: response.accessToken, accountId: response.account.homeAccountId })
    );
  }
}

export async function dialogFallback(callback) {
  // Attempt to acquire token silently if user is already signed in.
  if (homeAccountId !== null) {
    const result = await publicClientApp.acquireTokenSilent(loginRequest);
    if (result !== null && result.accessToken !== null) {
      const response = await callGetUserData(result.accessToken);
      callbackFunction(response);
    }
  } else {
    callbackFunction = callback;

    // We fall back to Dialog API for any error.
    const url = "/fallbackauthdialog.html";
    showLoginPopup(url);
  }
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
  // Uncomment to view message content in debugger, but don't deploy this way since it will expose the token.
  //console.log("Message received in processMessage: " + JSON.stringify(arg));

  let messageFromDialog = JSON.parse(arg.message);

  if (messageFromDialog.status === "success") {
    // We now have a valid access token.
    loginDialog.close();

    // Configure MSAL to use the signed-in account as the active account for future requests.
    const homeAccount = publicClientApp.getAccountByHomeId(messageFromDialog.accountId);
    if (homeAccount) {
      homeAccountId = messageFromDialog.accountId; // Track the account id for future silent token requests.
      publicClientApp.setActiveAccount(homeAccount);
    }

    const response = await callGetUserData(messageFromDialog.result);
    callbackFunction(response);
  } else if (messageFromDialog.error === undefined && messageFromDialog.result.errorCode === undefined) {
    // Need to pick the user to use to auth
  } else {
    // Something went wrong with authentication or the authorization of the web application.
    loginDialog.close();
    if (messageFromDialog.error) {
      showMessage(JSON.stringify(messageFromDialog.error.toString()));
    } else if (messageFromDialog.result) {
      showMessage(JSON.stringify(messageFromDialog.result.errorMessage.toString()));
    }
  }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
  var fullUrl = location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + url;

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function (result) {
    console.log("Dialog has initialized. Wiring up events");
    loginDialog = result.value;
    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
}
