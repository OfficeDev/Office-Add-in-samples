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
    pca: msalBrowser.IPublicClientApplication | undefined = undefined;

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
        if (this.pca === undefined) throw new Error("AccountManager is not initialized!");
        // Specify minimum scopes needed for the access token.
        const tokenRequest = {
            scopes: ["Files.Read"],
            loginHint: myloginHint
        }
        try {
            const userAccount = await this.pca.acquireTokenSilent(tokenRequest);
            return userAccount.accessToken;

        } catch (error) {
            // Acquire token silent failure. Send an interactive request via popup.
            try {
                const userAccount = await this.pca.acquireTokenPopup(tokenRequest);
                return userAccount.accessToken;
            } catch (popupError) {
                // Acquire token interactive failure.
                console.log(popupError);
                throw new Error("Unable to acquire access token: " + popupError)
            }
        }
    }

    /**
     * Uses MSAL and nested app authentication to get the user account from Office SSO.
     * This demonstrates how to work with user identity from the token.
     * 
     * @returns The user account data (identity).
     */
    async ssoGetUserIdentity() {
        if (this.pca === undefined) throw new Error("AccountManager is not initialized!");
        // Specify minimum scopes needed for the access token.
        const tokenRequest = {
            scopes: ["openid"],
            loginHint: myloginHint
        };
        // Acquire token silent failure. Send an interactive request via popup.
        try {
            const userAccount = await this.pca.acquireTokenPopup(tokenRequest);
            return userAccount;
        } catch (popupError) {
            // Acquire token interactive failure.
            console.log(popupError);
            throw new Error("Unable to acquire access token: " + popupError)
        }
    }
}