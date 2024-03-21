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
        $('#getUserFileNames').on("click", getUserFileNames);
    });
});

let retryGetAccessToken = 0;

async function getUserFileNames(options) {
    if (options === undefined) {
        options = { allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true };
    }
    clearMessage();

    // TODO 1: Get access token and call application server REST API.

    // TODO 2: Write the list of files to the document.

}

async function callRESTApi(relativeUrl, accessToken) {

    // TODO 3: Call the REST API on the application server.
   
}

function handleClientSideErrors(error) {

    //TODO 4: handle client side errors

    
}

function handleServerSideErrors(errorResponse) {

    //TODO 5: Check for admin consent error.

    //TODO 6: Check for additional claims request.

    //TODO 7: Check for expired token.

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