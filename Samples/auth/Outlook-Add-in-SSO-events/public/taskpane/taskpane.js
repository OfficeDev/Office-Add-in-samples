/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    
      document.getElementById("getProfileButton").onclick = getConsent;
    
  }
});

async function getConsent(){
  const accessToken = await OfficeRuntime.auth.getAccessToken({ allowConsentPrompt: true });
  document.getElementById("message-area").innerText = accessToken;
  console.log(accessToken);
}