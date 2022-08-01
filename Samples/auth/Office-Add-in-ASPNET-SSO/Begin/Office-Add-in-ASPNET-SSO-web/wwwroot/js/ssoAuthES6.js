// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// global to track if we are using SSO or the fallback auth.
// To test fallback auth, set authSSO = false.
let authSSO = true;

// If the add-in is running in Internet Explorer, the code must add support
// for Promises.
if (!window.Promise) {
    window.Promise = Office.Promise;
}

Office.onReady(function (info) {
    $(function () {
        $("#getFileNameListButton").on("click", getFileNameList);
    });
});


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
    // TODO 1: Initialize the client request.

    // TODO 2: Get the access token.

}



/**
 * Handles the click event for the Get File Name List button.
 * Requests a call to the ASP.NET Core server /api/filenames REST API that
 * gets up to 10 file names listed in the user's OneDrive.
 * When the call is completed, it will call the clientRequest.callbackRESTApiHandler.
 */
function getFileNameList() {
    clearMessage(); // Clear message log on task pane each time an API runs.

    // TODO 3: Create client request and call REST API.

}

/**
 * Handler for the returned response from the ASP.NET Core server API call to get file names.
 * Writes out the file names to the document.
 *
 * @param {*} response The list of file names.
 */
async function handleGetFileNameResponse(response) {
   // TODO 4: Pass response to writeFileNamesToOfficeDocument.
  //         Check for error and display success or error message.
}

/**
 * Calls the REST API on the server. Error handling will
 * switch to fallback auth if SSO fails.
 *
 * @param {*} clientRequest Contains information for calling an API on the server.
 */
async function callWebServer(clientRequest) {
    // TODO 5: Call REST API and check for errors.
    //         Get refreshed SSO token if current one expired.

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
   // TODO 7: Handle errors where the add-in should NOT invoke
    //         the alternative system of authorization.

    // TODO 8: Handle errors where the add-in should invoke
    //         the alternative system of authorization.
}

/**
 * Switches the client request to use MSAL.js auth (fallback) instead of SSO.
 * Once the new client request is created with MSAL.js access token, callWebServer is called
 * to continue attempting to call the REST API.
 * @param {*} clientRequest Contains information for calling an API on the server.
 */
function switchToFallbackAuth(clientRequest) {
    // TODO 9: Get a new client request to use MSAL.js.
}
