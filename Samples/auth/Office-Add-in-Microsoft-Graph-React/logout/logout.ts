/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { PublicClientApplication } from '@azure/msal-browser';

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {

    const msalInstance = new PublicClientApplication({
        auth: {
            clientId: 'fc19440a-334e-471e-af53-a1c1f53c9226',
            redirectUri: 'https://localhost:3000/logoutcomplete/logoutcomplete.html', 
            postLogoutRedirectUri: 'https://localhost:3000/logoutcomplete/logoutcomplete.html'
        }
    });

    msalInstance.logout();
  };
})();
