// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides functions to get ask the Office host to get an access token to the add-in
    and to pass that token to the server to get Microsoft Graph data. 
*/

// To support IE (which uses ES5), we have to create a Promise object for the global context.
if (!window.Promise) {
    window.Promise = Office.Promise;
}

Office.onReady(() => {
    $(document).ready(function () {
        $('#getUserFileNames').click(getUserFileNames);
    });
});

let retryGetAccessToken = 0;

async function getUserFileNames(options) {
    if (options === undefined) {
        options = { allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true };
    }
    clearMessage();
    let fileNameList = null;
    try {
        let accessToken = await Office.auth.getAccessToken(options);
        fileNameList = await callRESTApi("/api/files", accessToken);
    }
    catch (exception) {
        if (exception.code) {
            handleClientSideErrors(exception);
        }
        else {
            showMessage("EXCEPTION: " + exception);
        }
    }
    try {
        await writeFileNamesToOfficeDocument(fileNameList);
        showMessage("Your data has been added to the document.");
    } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin 
        // "Unable to add filenames to document."
        showMessage(error);
    }
}

async function callRESTApi(relativeUrl, accessToken) {
    try {
        let result = await $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
            dataType: "json",
            contentType: "application/json; charset=utf-8"
        });
        return result;
    } catch (error) {
        handleServerSideErrors(error);
    }
}

function handleClientSideErrors(error) {
    switch (error.code) {
        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one 
            // is logged into Office, then the first call of getAccessToken should pass the 
            // `allowSignInPrompt: true` option.
            showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.");
            break;
        case 13006:
            // Only seen in Office on the web.
            showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again.");
            break;
        case 13008:
            // Only seen in Office on the web.
            showMessage("Office is still working on the last operation. When it completes, try this operation again.");
            break;
        case 13010:
            // Only seen in Office on the web.
            showMessage("Follow the instructions to change your browser's zone configuration.");
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in by using MSAL authentication.
            showMessage("SSO failed. In these cases you should implement a falback to MSAL authentication.");
            break;
    }
}

function handleServerSideErrors(errorResponse) {
    // Check headers to see if admin has not consented.
    const header = errorResponse.getResponseHeader('WWW-Authenticate');
    if (header !== null && header.includes('proposedAction=\"consent\"')) {
        showMessage("MSAL ERROR: " + "Admin consent required. Be sure admin consent is granted on all scopes in the Azure app registration.");
        return;
    }

    // Check if Microsoft Graph requires an additional form of authentication. Have the Office host 
    // get a new token using the Claims string, which tells AAD to prompt the user for all 
    // required forms of authentication.
    const errorDetails = JSON.parse(errorResponse.responseJSON.value.details);
    if (errorDetails) {
        if (errorDetails.error.message.includes("AADSTS50076")) {
            const claims = errorDetails.message.Claims;
            const claimsAsString = JSON.stringify(claims);
            getUserFileNames({ authChallenge: claimsAsString });
            return;
        }
    }

    // Results from other errors (other than AADSTS50076) will have an ExceptionMessage property.
    const exceptionMessage = JSON.parse(errorResponse.responseText).ExceptionMessage;
    if (exceptionMessage) {
        // On rare occasions the access token is unexpired when Office validates it,
        // but expires by the time it is sent to Microsoft identity in the OBO flow. Microsoft identity will respond
        // with "The provided value for the 'assertion' is not valid. The assertion has expired."
        // Retry the call of getAccessToken (no more than once). This time Office will return a 
        // new unexpired access token.
        if ((exceptionMessage.includes("AADSTS500133"))
            && (retryGetAccessToken <= 0)) {
            retryGetAccessToken++;
            getGraphData();
            return;
        }
        else {
            showMessage("MSAL error from application server: " + JSON.stringify(exceptionMessage));
            return;
        }
    }
    // Default error handling if previous checks didn't apply.
    showMessage(errorResponse.responseJSON.value);
}

function showMessage(text) {
    const appendedText = $('#message-area').html() + text + "<br>---";
    $('.welcome-body').hide();
    $('#message-area').show();
    $('#message-area').html(appendedText);
}

function clearMessage() {
    $('.welcome-body').hide();
    $('#message-area').show();
    $('#message-area').html("---<br>");
}