/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

// Replace with the client ID from your fallback auth app registration.
const fallbackClientID = "$fallback_application_GUID_here$";

  // If the add-in is running in Internet Explorer, the code must add support 
 // for Promises.
if (!window.Promise) {
    window.Promise = Office.Promise;
}

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit: 
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
 const loginRequest = {
   scopes: ["access_as_user"],
  };



    Office.initialize = function () { 
        if (Office.context.ui.messageParent)
        {
            debugger;
            userAgentApp.handleRedirectCallback(authCallback);

            // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
            // stored login data in localStorage. So a direct call of acquireTokenRedirect 
            // causes the error "User login is required". Once the user is logged in successfully
            // the first time, msal data in localStorage will prevent this error from ever hap-
            // pening again; but the error must be blocked here, so that the user can login 
            // successfully the first time. To do that, call loginRedirect first instead of 
            // acquireTokenRedirect.
            if (localStorage.getItem("loggedIn") === "yes") {
                userAgentApp.acquireTokenRedirect(loginRequest);
            }
            else {
                // This will login the user and then the (response.tokenType === "id_token")
                // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
                // and then the dialog is redirected back to this script, so the 
                // acquireTokenRedirect above runs.
                userAgentApp.loginRedirect(loginRequest);
            }
        }
     };

    // const msalConfig = {
    //     auth: {
    //         clientId: fallbackClientID,
    //         //authority: "https://login.microsoftonline.com/common", 
    //         authority: "https://login.microsoftonline.com/"+ fallbackClientID, 
    //         redirectURI: "https://localhost:44355/dialog", 
    //         navigateToLoginRequestUrl: false,
    //         response_type: "access_token"
    //     },
    //     cache: {
    //         cacheLocation: 'localStorage', // Needed to avoid "User login is required" error.
    //         storeAuthStateInCookie: true  // Recommended to avoid certain IE/Edge issues.
    //     }
    // };

    const userAgentApp = new Msal.UserAgentApplication(msalConfig);

    function authCallback(error, response) {
        if (error) {
            console.log(error);
            Office.context.ui.messageParent(JSON.stringify({ status: 'failure', result : error }));
        } else {
            if (response.tokenType === "id_token") {
                console.log(response.idToken.rawIdToken);
                localStorage.setItem("loggedIn", "yes");
            } else {
                console.log("token type is:" + response.tokenType);
                Office.context.ui.messageParent( JSON.stringify({ status: 'success', result : response.accessToken }) );               
            }        
        }
    }




    