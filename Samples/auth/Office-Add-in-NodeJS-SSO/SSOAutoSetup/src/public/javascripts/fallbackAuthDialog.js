/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

 // If the add-in is running in Internet Explorer, the code must add support 
 // for Promises.
 if (!window.Promise) {
    window.Promise = Office.Promise;
}

Office.initialize = function () { 
    if (Office.context.ui.messageParent)
    {
        userAgentApp.handleRedirectCallback(authCallback);

        // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
        // stored login data in localStorage. So a direct call of acquireTokenRedirect 
        // causes the error "User login is required". Once the user is logged in successfully
        // the first time, msal data in localStorage will prevent this error from ever hap-
        // pening again; but the error must be blocked here, so that the user can login 
        // successfully the first time. To do that, call loginRedirect first instead of 
        // acquireTokenRedirect.
        if (localStorage.getItem("loggedIn") === "yes") {
            userAgentApp.acquireTokenRedirect(requestObj);
        }
        else {
            // This will login the user and then the (response.tokenType === "id_token")
            // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
            // and then the dialog is redirected back to this script, so the 
            // acquireTokenRedirect above runs.
            userAgentApp.loginRedirect(requestObj);
        }
    }
 };

const msalConfig = {
    auth: {
        clientId: "{application GUID here}", //This is your client ID
        authority: "https://login.microsoftonline.com/common", 
        redirectURI: "https://localhost:44355/dialog.html", 
        navigateToLoginRequestUrl: false,
        response_type: "access_token"
    },
    cache: {
        cacheLocation: 'localStorage', // Needed to avoid "User login is required" error.
        storeAuthStateInCookie: true  // Recommended to avoid certain IE/Edge issues.
    }
};

var requestObj = {
    scopes: ["https://graph.microsoft.com/Files.Read.All"]
};

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