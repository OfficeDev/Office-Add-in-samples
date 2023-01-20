// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new msal.PublicClientApplication(msalConfig);

let username = '';

/**
 * This method adds an event callback function to the MSAL object
 * to handle the response from redirect flow. For more information, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/events.md
 */
myMSALObj.addEventCallback((event) => {
    if (
        (event.eventType === msal.EventType.LOGIN_SUCCESS ||
        event.eventType === msal.EventType.ACQUIRE_TOKEN_SUCCESS) &&
        event.payload.account
    ) {
        const account = event.payload.account;
        myMSALObj.setActiveAccount(account);
    }

    if (event.eventType === msal.EventType.LOGOUT_SUCCESS) {
        if (myMSALObj.getAllAccounts().length > 0) {
            myMSALObj.setActiveAccount(myMSALObj.getAllAccounts()[0]);
        }
    }
});

function selectAccount() {
    /**
     * See here for more info on account retrieval:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */
    const currentAccounts = myMSALObj.getAllAccounts();
    if (currentAccounts === null) {
        return;
    } else if (currentAccounts.length >= 1) {
        // Add choose account code here
        username = myMSALObj.getActiveAccount().username;
        showWelcomeMessage(username, currentAccounts);
    }
}

async function addAnotherAccount(event) {
    if (event.target.innerHTML.includes('@')) {
        const username = event.target.innerHTML;
        const account = myMSALObj.getAllAccounts().find((account) => account.username === username);
        const activeAccount = myMSALObj.getActiveAccount();
        if (account && activeAccount.homeAccountId != account.homeAccountId) {
            try {
                myMSALObj.setActiveAccount(account);
                let res = await myMSALObj.ssoSilent({
                    ...loginRequest,
                    account: account,
                });
                closeModal();
                handleResponse(res);
                window.location.reload();
            } catch (error) {
                if (error instanceof msal.InteractionRequiredAuthError) {
                    let res = await myMSALObj.loginPopup({
                        ...loginRequest,
                        prompt: 'login',
                    });
                    handleResponse(res);
                    window.location.reload();
                }
            }
        } else {
            closeModal();
        }
    } else {
        try {
            myMSALObj.setActiveAccount(null);
            const res = await myMSALObj.loginPopup({
                ...loginRequest,
                prompt: 'login',
            });
            handleResponse(res);
            closeModal();
            window.location.reload();
        } catch (error) {
            console.log(error);
        }
    }
}

function handleResponse(response) {
    /**
     * To see the full list of response object properties, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#response
     */

    if (response !== null) {
        const accounts = myMSALObj.getAllAccounts();
        username = response.account.username;
        showWelcomeMessage(username, accounts);
    } else {
        selectAccount();
    }
}

function signIn() {
    /**
     * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
     */

    myMSALObj
        .loginPopup(loginRequest)
        .then(handleResponse)
        .catch((error) => {
            console.error(error);
        });
}

function signOut() {
    /**
     * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
     */
    const account = myMSALObj.getAccountByUsername(username);
    const logoutRequest = {
        account: account,
        mainWindowRedirectUri: '/',
    };
    clearStorage(account);
    myMSALObj.logoutPopup(logoutRequest).catch((error) => {
        console.log(error);
    });
}

function seeProfile() {
    callGraph(
        username,
        graphConfig.graphMeEndpoint.scopes,
        graphConfig.graphMeEndpoint.uri,
        msal.InteractionType.Popup,
        myMSALObj
    );
}

function readContacts() {
    callGraph(
        username,
        graphConfig.graphContactsEndpoint.scopes,
        graphConfig.graphContactsEndpoint.uri,
        msal.InteractionType.Popup,
        myMSALObj
    );
}

selectAccount();
