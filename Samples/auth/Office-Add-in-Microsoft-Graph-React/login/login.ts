/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { PublicClientApplication } from "@azure/msal-browser";

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {

    const msalInstance = new PublicClientApplication({
        auth: {
          clientId: 'fc19440a-334e-471e-af53-a1c1f53c9226',
          authority: 'https://login.microsoftonline.com/common',
          redirectUri: 'https://localhost:3000/login/login.html' // Must be registered as "spa" type
        },
        cache: {
          cacheLocation: 'localStorage', // needed to avoid "login required" error
          storeAuthStateInCookie: true   // recommended to avoid certain IE/Edge issues
        }
      });

    // handleRedirectPromise should be invoked on every page load
    msalInstance.handleRedirectPromise()
        .then((response) => {
            // If response is non-null, it means page is returning from AAD with a successful response
            if (response) {
                Office.context.ui.messageParent( JSON.stringify({ status: 'success', result : response.accessToken }) );
            } else {
                // Otherwise, invoke login
                msalInstance.loginRedirect({
                    scopes: ['user.read', 'files.read.all']
                });
            }
        })
        .catch((error) => {
            const errorData: string = `errorMessage: ${error.errorCode}
                                   message: ${error.errorMessage}
                                   errorCode: ${error.stack}`;
            Office.context.ui.messageParent( JSON.stringify({ status: 'failure', result: errorData }));
        });
  };
})();
