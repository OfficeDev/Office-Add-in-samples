// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// global to track if we are using SSO or the fallback auth.
// To test fallback auth, set authSSO = false.
let authSSO = false;

// If the add-in is running in Internet Explorer, the code must add support
// for Promises.
if (!window.Promise) {
    window.Promise = Office.Promise;
}

Office.onReady(function (info) {
    $(function () {
        $("#getFileNameListButton").on("click", getFileNameList);
        $("#signInButton").on("click", dialogFallback);
    });
});

/**
 * Handles the click event for the Get File Name List button.
 * Requests a call to the ASP.NET Core server /api/filenames REST API that
 * gets up to 10 file names listed in the user's OneDrive.
 * When the call is completed, it will call the clientRequest.callbackRESTApiHandler.
 */
function getFileNameList() {
    clearMessage(); // Clear message log on task pane each time an API runs.
    createRequest(
        "GET",
        "/api/filenames",
        handleGetFileNameResponse,
        async (clientRequest) => {
            await callWebServer(clientRequest);
        }
    );
}

/**
 * Handler for the returned response from the ASP.NET Core server API call to get file names.
 * Writes out the file names to the document.
 *
 * @param {*} response The list of file names.
 */
async function handleGetFileNameResponse(response) {
    if (response !== null) {
        try {
            await writeFileNamesToOfficeDocument(response);
            showMessage("Your OneDrive filenames are added to the document.");
        } catch (error) {
            // The error from writeFileNamesToOfficeDocument will begin
            // "Unable to add filenames to document."
            showMessage(error);
        }
    } else
        showMessage("A null response was returned to handleGetFileNameResponse.");
}

/**
 * Calls the REST API on the server. Error handling will
 * switch to fallback auth if SSO fails.
 *
 * @param {*} clientRequest Contains information for calling an API on the server.
 */
async function callWebServer(clientRequest) {    
    try {
        const data = await $.ajax({
            type: clientRequest.verb,
            url: clientRequest.url,
            headers: {"Authorization": "Bearer " + clientRequest.accessToken},            
            cache: false
        });
        clientRequest.callbackRESTApiHandler(data);
    } catch (error) {
        // Check for expired token. Refresh and retry the call if it expired.
        if (error.getResponseHeader !== undefined) {
            const responseHeader = error.getResponseHeader("www-authenticate");
            if (responseHeader !== null && responseHeader.includes("The token expired") && authSSO) {
                try {
                    clientRequest.accessToken = await Office.auth.getAccessToken(clientRequest.authOptions);
                    const data = await $.ajax({
                        type: clientRequest.verb,
                        url: clientRequest.url,
                        headers: { Authorization: "Bearer " + clientRequest.accessToken },
                        xhrFields: {
                            withCredentials: true
                        },
                        cache: false
                    });
                    clientRequest.callbackRESTApiHandler(data);
                } catch (error) {
                    showMessage(error.responseText);
                    switchToFallbackAuth(clientRequest);
                    return;
                }
            }
        }

        // Check for a Microsoft Graph API call error. which is returned as bad request (403)
        if (error.status === 403) {
            showMessage(error.reponseText);
            return;
        }

        // For all other error scenarios, display the message and use fallback auth.
        showMessage(
            "Unknown error from web server: " +
            JSON.stringify(error.responseJSON.errorDetails)
        );
        if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    }
}

/**
 * Switches the client request to use MSAL.js auth (fallback) instead of SSO.
 * Once the new client request is created with MSAL.js access token, callWebServer is called
 * to continue attempting to call the REST API.
 * @param {*} clientRequest Contains information for calling an API on the server.
 */
function switchToFallbackAuth(clientRequest) {
    // Guard against accidental call to this function when fallback is already in use.
    if (authSSO === false) return;

    showMessage("Switching from SSO to fallback auth.");
    authSSO = false;
    // Create a new request for fallback auth.
    createRequest(
        clientRequest.verb,
        clientRequest.url,
        clientRequest.callbackRESTApiHandler,
        async (fallbackRequest) => {
            // Hand off to call using fallback auth.
            await callWebServer(fallbackRequest);
        }
    );
}

/**
 * Creates a client request object with:
 * authOptions - Auth configuration parameters for SSO.
 * authSSO - true if using SSO, otherwise false.
 * verb - REST API verb such as GET, POST...
 * accessToken - The access token to the ASP.NET Core server.
 * url - The URL of the REST API to call on the ASP.NET Core server.
 * callbackRESTApiHandler - The function to pass the results of the REST API call.
 * callbackFunction - the function to pass the client request to when ready.
 *
 * Note that when the client request is created it will be passed to the callbackFunction. This is used because
 * we may need to pop up a dialog to sign in the user, which uses a callback approach.
 *
 * @param {*} callbackFunction The function to pass the client request to when ready.
 */
async function createRequest(verb, url, restApiCallback, callbackFunction) {
    const clientRequest = {
        authOptions: {
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: true,
        },
        authSSO: authSSO,
        verb: verb,
        accessToken: null,
        url: url,
        callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };

    if (authSSO) {
        try {
            // Get access token from Office SSO.
            clientRequest.accessToken = await Office.auth.getAccessToken(clientRequest.authOptions);
            callbackFunction(clientRequest);
        } catch (error) {
            // handle the SSO error which will inform us if we need to switch to fallback auth.
            let fallbackRequired = handleSSOErrors(error);
            if (fallbackRequired) switchToFallbackAuth(clientRequest);
        }
    } else {
        // Use fallback auth to get access token.
        dialogFallback(clientRequest);
    }
}

/**
 * Handles any error returned from getAccessToken. The numbered errors are typically user actions
 * that don't require fallback auth. The text shown for each error indicates next steps
 * you should take. For default (all other errors), the sample returns true
 * so that the caller is informed to use fallback auth.
 *
 * @param {*} err The error to process.
 * @returns true if SSO error could not be handled, and fallback auth is required; otherwise, false.
 */
function handleSSOErrors(err) {
    let fallbackRequired = false;
    switch (err.code) {
        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one
            // is logged into Office, then the first call of getAccessToken should pass the
            // `allowSignInPrompt: true` option. Since this sample does that, you should not see
            // this error.
            showMessage(
                "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
            );
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage(
                "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
            );
            break;
        case 13006:
            // Only seen in Office on the web.
            showMessage(
                "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
            );
            break;
        case 13008:
            // Only seen in Office on the web.
            showMessage(
                "Office is still working on the last operation. When it completes, try this operation again."
            );
            break;
        case 13010:
            // Only seen in Office on the web.
            showMessage(
                "Follow the instructions to change your browser's zone configuration."
            );
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in.
            fallbackRequired = true;
            break;
    }
    return fallbackRequired;
}
