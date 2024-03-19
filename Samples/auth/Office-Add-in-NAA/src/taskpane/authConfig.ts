// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

import * as msalBrowser from "@azure/msal-browser";

export { AccountManager }

const applicationId = "Enter_the_Application_Id_Here";
const myloginHint = "Enter_the_Login_Hint_Here";

const msalConfig = {
    auth: {
        clientId: applicationId,
        authority: "https://login.microsoftonline.com/common",
        supportsNestedAppAuth: true
    }
}

// Encapsulate functions for getting user account and token information.
class AccountManager {
    pca = undefined;

    // Initialize MSAL public client application.
    async initialize() {
        this.pca = await msalBrowser.PublicClientNext.createPublicClientApplication(msalConfig);
    }

    /**
     * Uses MSAL and nested app authentication to get an access token through Office SSO.
     * Call this function any time you need an access token for Microsoft Graph.
     * 
     * @returns An access token for calling Microsoft Graph APIs.
     */
    async ssoGetToken() {
        const tokenRequest = {
            scopes: ["Files.Read"],
            loginHint: myloginHint
        };
        try {
            const userAccount = await this.pca.ssoSilent(tokenRequest);
            return userAccount.accessToken;
        } catch (error) {
            console.log(error);
            let resultatpu = this.pca.acquireTokenPopup(tokenRequest);
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
    async ssoGetUserIdentity() {
        const tokenRequest = {
            scopes: ["openid"],
            loginHint: myloginHint
        };
        try {
            const userAccount = await this.pca.ssoSilent(tokenRequest);
            return userAccount;
        } catch (error) {
            console.log(error);
            let resultatpu = this.pca.acquireTokenPopup(tokenRequest);
            console.log("result: " + resultatpu);
            throw (error); //rethrow
        }
    }
}