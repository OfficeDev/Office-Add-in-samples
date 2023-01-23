// Create the main myMSALObj instance

//const { json } = require("express");

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
        const account = myMSALObj
            .getAllAccounts()
            .find((account) => account.username === username);
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

function openInExcel() {
    // create test table data to pass to Azure function.
    const tableData = {
        rows: [
            {
                columns: [
                    { value: 'ID' },
                    { value: 'Name' },
                    { value: 'Qtr1' },
                    { value: 'Qtr2' },
                    { value: 'Qtr3' },
                    { value: 'Qtr4' },
                ],
            },
            {
                columns: [
                    { value: '1' },
                    { value: 'Frames' },
                    { value: '5000' },
                    { value: '7000' },
                    { value: '6544' },
                    { value: '4377' },
                ],
            },
            {
                columns: [
                    { value: '2' },
                    { value: 'Saddles' },
                    { value: '400' },
                    { value: '323' },
                    { value: '276' },
                    { value: '651' },
                ],
            },
        ],
    };

    //    var products = new List<Product>()
    //{ new Product {ID=1, Name="Frames", Qtr1=5000, Qtr2=7000, Qtr3=6544, Qtr4=4377},
    //new Product {ID=2, Name="Saddles", Qtr1=400, Qtr2=323, Qtr3=276, Qtr4=651},
    //new Product {ID=3, Name="Brake levers", Qtr1=12000, Qtr2=8766, Qtr3=8456, Qtr4=9812},
    //new Product {ID=4, Name="Chains", Qtr1=1550, Qtr2=1088, Qtr3=692, Qtr4=853},
    //new Product {ID=5, Name="Mirrors", Qtr1=225, Qtr2=600, Qtr3=923, Qtr4=544},
    //new Product {ID=5, Name="Spokes", Qtr1=6005, Qtr2=7634, Qtr3=4589, Qtr4=8765}
    //}
    //const bodyEncoded = encodeURIComponent(JSON.stringify(tableData));
    const bodyEncoded = JSON.stringify(tableData);
    // get spreadsheet
    //http://localhost:7071/api/Function1
    fetch('http://localhost:7071/api/Function1', {
        headers: {
            'Content-Type': 'application/octet-stream',
        },
        method: 'POST',
        body: bodyEncoded,
    })
        .then((response) => response.blob())
        .then((blob) => {
            console.log(blob);
            uploadFile('openinexcel', 'spreadsheet.xlsx', blob);
        });
}

function uploadFile(folderName, fileName, data) {
    const uri =
        'https://graph.microsoft.com/v1.0/me/drive/root:/' +
        folderName +
        '/' +
        fileName +
        ':/content';

    callGraph(
        username,
        graphConfig.graphFilesEndpoint.scopes,
        uri,
        msal.InteractionType.Popup,
        myMSALObj,
        data
    );
}

selectAccount();
