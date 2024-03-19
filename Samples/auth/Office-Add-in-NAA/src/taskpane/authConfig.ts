// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

import * as msalBrowser from "@azure/msal-browser";
import { getGraphData } from "./msgraph-helper";

export { ssoGetToken, ssoGetUserIdentity }

//msal config - using dev tenant registered app ID. 
const msalConfig = {
    auth: {
        clientId: "57e00eca-d992-4e1c-bef6-a238cd0236c4",
        authority: "https://login.microsoftonline.com/common",
        supportsNestedAppAuth: true
    }
}
const myloginhint = "davechuatest3.onmicrosoft.com";

// Initialize MSAL public client application.
let pca = undefined;
msalBrowser.PublicClientNext.createPublicClientApplication(msalConfig).then((result) => {
    pca = result;
});

/**
 * Uses MSAL and nested app authentication to get an access token through Office SSO.
 * Call this function any time you need an access token for Microsoft Graph.
 * 
 * @returns An access token for calling Microsoft Graph APIs.
 */
async function ssoGetToken() {
    //const activeAccount = pca.getActiveAccount();  
    const tokenRequest = {
        scopes: ["User.Read", "Files.Read", "openid", "profile"],
        loginhint: myloginhint
    };
    try {
        const userAccount = await pca.ssoSilent(tokenRequest);
        return userAccount.accessToken;
    } catch (error) {
        console.log(error);
        let resultatpu = pca.acquireTokenPopup(tokenRequest);
        console.log("result: " + resultatpu);
        throw (error); //rethrow
    }
}

/**
 * Uses MSAL and nested app authentication to get the user account from Office SSO.
 * This demonstrates how to work with user identity from the token.
 * 
 * @returns The user account data (identity).
 */
async function ssoGetUserIdentity() {
    const tokenRequest = {
        scopes: [ "openid" ],
        loginhint: myloginhint
    };
    try {
        const userAccount = await pca.ssoSilent(tokenRequest);
        return userAccount;
    } catch (error) {
        console.log(error);
        let resultatpu = pca.acquireTokenPopup(tokenRequest);
        console.log("result: " + resultatpu);
        throw (error); //rethrow
    }
}