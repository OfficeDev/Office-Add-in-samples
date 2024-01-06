/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      document.getElementById("getProfileButton").onclick = signInAndConsent;
  }
});

/**
 * Gets an access token for the user. This will force sign in and consent
 * if the user has not already signed in.
 */
async function signInAndConsent(){
  try {
    const accessToken = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false 
    });
    document.getElementById("message-area").innerText = "Sign in successful.";
  } catch (exception){
    document.getElementById("message-area").innerText = "There was an error signing in: " + exception.message;
  }
}