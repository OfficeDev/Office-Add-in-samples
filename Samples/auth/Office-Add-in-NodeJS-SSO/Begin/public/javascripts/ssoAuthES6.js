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
 * When the call is completed, it will call the clientRequest.callbackRESTApiHandler.
 */
async function getFileNameList() {
  clearMessage(); // Clear message log on task pane each time an API runs.

  // TODO 1: Call server API, then write filenames to the Office document.

}

  /**
 * Handles any error returned from getAccessToken. The numbered errors are typically user actions
 * that don't require fallback auth. The text shown for each error indicates next steps
 * you should take. For default (all other errors), the sample returns true
 * so that the caller is informed to use fallback auth.
 * @param {*} error The error returned by Office.auth.getAccessToken.
 * @returns access token when falling back to MSAL auth; otherwise, null.
 */
   function handleSSOErrors(err) {

    // TODO 2: Handle errors where the add-in should NOT invoke 
    //         the alternative system of authorization.

    // TODO 3: Handle errors where the add-in should invoke 
    //         the alternative system of authorization.

   }

/**
 * Call our server REST API and return the response JSON.
 * @param {*} method Which HTTP method to use.
 * @param {*} path The URL path of the server REST API.
 * @param {*} retryRequest Indicates if this is a retry of the call.
 * @returns The response JSON from the server API.
 */
async function callServerAPI(method, path, retryRequest = false) {
    
    // TODO 4: Get access token, then make fetch call to the server REST API.

    // TODO 5: Check for expired SSO token.

    // TODO 6: Check for Microsoft Graph errors.

    // TODO 7: Check for other errors.

  }

/**
 * Gets an access token for the user. Will use SSO if available, otherwise will use MSAL.
 */
async function getAccessToken(authSSO) {

  // TODO 8: Attempt to get an access token using SSO.

  // TODO 9: Attempt to get an access token using MSAL.
}