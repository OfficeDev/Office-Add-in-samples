/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { PublicClientApplication } from "@azure/msal-browser";

Office.onReady(async () => {
  const pca = new PublicClientApplication({
    auth: {
      clientId: 'YOUR APP ID HERE',
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: 'https://localhost:3000/login/login.html' // Must be registered as "spa" type
    },
    cache: {
      cacheLocation: 'localStorage', // needed to avoid "login required" error
      storeAuthStateInCookie: true   // recommended to avoid certain IE/Edge issues
    }
  });
  await pca.initialize();

  try {
    // handleRedirectPromise should be invoked on every page load
    const response = await pca.handleRedirectPromise();
    if (response) {
      Office.context.ui.messageParent(JSON.stringify({ status: 'success', token: response.accessToken, userName: response.account.username}));
    } else {
      // Problem occurred, so invoke login
      pca.loginRedirect({
        scopes: ['user.read', 'files.read.all']
      });
    }
  } catch (error) {
    const errorData = {
      errorMessage: error.errorCode,
      message: error.errorMessage,
      errorCode: error.stack
    };
    Office.context.ui.messageParent(JSON.stringify({ status: 'failure', result: errorData }));
  }
});

