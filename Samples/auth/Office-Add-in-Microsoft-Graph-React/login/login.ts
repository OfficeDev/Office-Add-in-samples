/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as msal from 'msal';

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {

    const config: msal.Configuration = {
      auth: {
        clientId: 'THE APPLICATION (CLIENT) ID GOES HERE',
        authority: 'https://login.microsoftonline.com/common',
        redirectUri: 'https://localhost:3000/login/login.html'
      },
      cache: {
        cacheLocation: 'localStorage', // needed to avoid "login required" error
        storeAuthStateInCookie: true   // recommended to avoid certain IE/Edge issues
      }
    };

    const userAgentApp: msal.UserAgentApplication = new msal.UserAgentApplication(config);

    const authCallback = (error: msal.AuthError, response: msal.AuthResponse) => {

      if (!error) {
        if (response.tokenType === 'id_token') {
          /*
            If the O365 tenancy is configured to require two-factor authentication, AAD
            sends an ID token and *REDIRECTS BACK TO THIS PAGE* immediately after the user
            provides a password, but before the user has been prompted for the 2nd factor.
            So, this immediately invoked function expression (IIFE) runs again. When
            acquireTokenRedirect runs a second time, AAD prompts the user for the 2nd factor
            and then returns the access token. This happens so fast that the user gets the
            2nd factor prompt immediately after providing a password. But this code needs to
            test for the case when only an ID token is returned and DO NOTHING, so that the
            IIFE runs again.
          */
        }
        else {
          // The tokenType is access_token, so send success message and token.
          Office.context.ui.messageParent( JSON.stringify({ status: 'success', result : response.accessToken }) );
        }
      }
      else {
        const errorData: string = `errorMessage: ${error.errorCode}
                                   message: ${error.errorMessage}
                                   errorCode: ${error.stack}`;
        Office.context.ui.messageParent( JSON.stringify({ status: 'failure', result: errorData }));
      }
    };

    userAgentApp.handleRedirectCallback(authCallback);

    const request: msal.AuthenticationParameters = {
      scopes: ['user.read', 'files.read.all'],
    };

    userAgentApp.acquireTokenRedirect(request);
  };
})();
