/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { PublicClientApplication } from '@azure/msal-browser';

let pca;

Office.onReady(async () => {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived,
    onMessageFromParent);
  pca = new PublicClientApplication({
    auth: {
      clientId: 'YOUR APP ID HERE',
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: `${window.location.origin}/login/login.html` // Must be registered as "spa" type.
    },
    cache: {
      cacheLocation: 'localStorage' // Needed to avoid a "login required" error.
    }
  });
  await pca.initialize();
});

async function onMessageFromParent(arg) {
  const messageFromParent = JSON.parse(arg.message);
  const currentAccount = pca.getAccountByHomeId(messageFromParent.userName);
  // You can select which account application should sign out.
  const logoutRequest = {
    account: currentAccount,
    postLogoutRedirectUri: `${window.location.origin}/logoutcomplete/logoutcomplete.html`,
  };
  await pca.logoutRedirect(logoutRequest);
  const messageObject = { messageType: "dialogClosed" };
  const jsonMessage = JSON.stringify(messageObject);
  Office.context.ui.messageParent(jsonMessage);
}
