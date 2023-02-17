// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// global to track if we are using SSO or the fallback auth.
// To test fallback auth, set authSSO = false.
let authSSO = true;

// Default SSO options when calling getAccessToken.
const ssoOptions = {
    allowSignInPrompt: true,
    allowConsentPrompt: true,
    forMSGraphAccess: true,
};

// If the add-in is running in Internet Explorer, the code must add support
// for Promises.
if (!window.Promise) {
    window.Promise = Office.Promise;
}

Office.onReady(() => {
    document
        .getElementById('getFileNameListButton')
        .addEventListener('click', getFileNameList);
});

/**
 * Handles the click event for the Get File Name List button.
 * Requests a call to the middle-tier server /getuserfilenames that
 * gets up to 10 file names listed in the user's OneDrive.
 */
async function getFileNameList() {
    clearMessage(); // Clear message log on task pane each time an API runs.

    try {
        const jsonResponse = await callWebServerAPI('GET', '/getuserfilenames');
        if (jsonResponse === null) {
            return; // When null is returned a message was already shown to the user prompting to try again with additional information.
        }
        await writeFileNamesToOfficeDocument(jsonResponse);
        showMessage('Your OneDrive filenames are added to the document.');
    } catch (error) {
        console.log(error.message);
        showMessage(error.message);
    }
}

/**
 * Call our server REST API and return the response JSON.
 * @param {*} method Which HTTP method to use.
 * @param {*} path The URL path of the server REST API.
 * @param {*} retryRequest Indicates if this is a retry of the call.
 * @returns The response JSON from the server API.
 */
async function callWebServerAPI(method, path, retryRequest = false) {
    const accessToken = await getAccessToken(authSSO);
    if (accessToken === null) {
        return null;
    }
    const response = await fetch(path, {
        method: method,
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + accessToken,
        },
    });

    // Check for success condition: HTTP status code 2xx.
    if (response.ok) {
        return response.json();
    }

    // Check for fail condition: Is SSO token expired? If so, retry the call which will get a refreshed token.
    const jsonBody = await response.json();
    if (
        authSSO === true &&
        jsonBody != null &&
        jsonBody.type === 'TokenExpiredError'
    ) {
        if (!retryRequest) {
            return callWebServerAPI(method, path, true); // Try the call again. The underlying call to Office JS getAccessToken will refresh the token.
        } else {
            // Indicates a second call to retry and refresh the token failed.
            authSSO = false;
            return callWebServerAPI(method, path, true); // Try the call again, but now using MSAL fallback auth.
        }
    }

    // Check for fail condition: Did we get a Microsoft Graph API error, which is returned as bad request (403)?
    if (response.status === 403 && jsonBody.type === 'Microsoft Graph') {
        throw new Error('Microsoft Graph error: ' + jsonBody.errorDetails);
    }

    // Handle other errors.
    throw new Error(
        'Unknown error from web server: ' + JSON.stringify(jsonBody)
    );
}

/**
 * Gets an access token for the user. Will use SSO if available, otherwise will use MSAL.
 */
async function getAccessToken(authSSO) {
    if (authSSO) {
        try {
            // Get the access token from Office host using SSO.
            // Note that Office.auth.getAccessToken modifies the options parameter. Create a copy of the object
            // to avoid modifying the original object.
            const options = JSON.parse(JSON.stringify(ssoOptions));
            const token = await Office.auth.getAccessToken(options);
            return token;
        } catch (error) {
            console.log(error.message);
            return handleSSOErrors(error);
        }
    } else {
        // Get access token through MSAL fallback.
        try {
            const accessToken = await getAccessTokenMSAL();
            return accessToken;
        } catch (error) {
            console.log(error);
            throw new Error(
                'Cannot get access token. Both SSO and fallback auth failed. ' +
                    error
            );
        }
    }
}

/**
* Handles any error returned from getAccessToken. The numbered errors are typically user actions
 * that don't require fallback auth. The text shown for each error indicates next steps
 * you should take. For default (all other errors), the sample returns true
 * so that the caller is informed to use fallback auth.
 * @param {*} error The error returned by Office.auth.getAccessToken.
 * @returns access token when falling back to MSAL auth; otherwise, null.
 */
async function handleSSOErrors(error) {
    switch (error.code) {
        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one
            // is logged into Office, then the first call of getAccessToken should pass the
            // `allowSignInPrompt: true` option. Since this sample does that, you should not see
            // this error.
            showMessage(
                'No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.'
            );
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage(
                'You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.'
            );
            break;
        case 13006:
            // Only seen in Office on the web.
            showMessage(
                'Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again.'
            );
            break;
        case 13008:
            // Only seen in Office on the web.
            showMessage(
                'Office is still working on the last operation. When it completes, try this operation again.'
            );
            break;
        case 13010:
            // Only seen in Office on the web.
            showMessage(
                "Follow the instructions to change your browser's zone configuration."
            );
            break;
        default: //recursive call.
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to MSAL sign-in.
            showMessage('SSO failed. Trying fallback auth.');
            authSSO = false;
            return getAccessToken(false);
    }
    return null; // Return null for errors that show a message to the user.
}
