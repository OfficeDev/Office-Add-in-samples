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
 * Handles the click event for the Get File Name List button.
 * Requests a call to the middle-tier server /getuserfilenames that
 * gets up to 10 file names listed in the user's OneDrive.
 * When the call is completed, it will call the clientRequest.callbackRESTApiHandler.
 */
function getFileNameList() {
  clearMessage(); // Clear message log on task pane each time an API runs.
  createRequest("/getuserfilenames",handleGetFileNameResponse, async (clientRequest) => {
    await callWebServer(clientRequest);
  });
}

/**
 * Handler for the returned response from the middle-tier server API call to get file names.
 * Writes out the file names to the document.
 *
 * @param {*} response The list of file names.
 */
async function handleGetFileNameResponse(response) {
  if (response != null) {
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
 * Calls the REST API on the middle-tier server. Error handling will
 * switch to fallback auth if SSO fails.
 *
 * @param {*} clientRequest Contains information for calling an API on the middle-tier server.
 */
async function callWebServer(clientRequest) {
  try {
    await ajaxCallToRESTApi(clientRequest);
  } catch (error) {
    if (error.statusText === "Internal Server Error") {
      const isTokenExpired = handleWebServerErrors(error);
      if (isTokenExpired && clientRequest.authSSO) {
        try {
          clientRequest.accessToken = await getAccessTokenFromSSO(clientRequest.authOptions);
          await ajaxCallToRESTApi(clientRequest);
        } catch {
          // If still an error go to fallback.
          switchToFallbackAuth(clientRequest);
          return;
        }
      } else if (error.statusText === "Missing access_as_user") {
        showMessage("Error: Access token is missing the access_as_user scope.");
      } else {
        // For unhandled errors using SSO, switch to fallback.
        if (clientRequest.authSSO) {
          switchToFallbackAuth(clientRequest);
        } else {
          console.log(JSON.stringify(error)); // Log any errors.
          showMessage(error.responseText);
        }
      }
    } else {
      console.log(JSON.stringify(error)); // Log any errors.
          showMessage(error.responseText);
    }
  }
}

/**
 * Makes the AJAX call to the REST API in the middle-tier server.
 * Note that any errors are thrown to the caller to handle.
 * @param {} clientRequest Contains information for calling an API on the middle-tier server.
 */
async function ajaxCallToRESTApi(clientRequest) {
  try {
    await $.ajax({
      type: "GET",
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
      success: function (data) {
        result = data;
        // Send result to the callback handler.
        clientRequest.callbackRESTApiHandler(result);
      },
    });
  } catch (error) {
    // This function explicitly requires the caller to handle any errors
    throw error;
  }
}

/**
 * Switches the client request to use MSAL auth (fallback) instead of SSO. 
 * Once the new client request is created with MSAL access token, callWebServer is called
 * to continue attempting to call the REST API.
 * @param {*} clientRequest Contains information for calling an API on the middle-tier server.
 */
function switchToFallbackAuth(clientRequest) {
  showMessage("Switching from SSO to fallback auth.");
  authSSO = false;
  // Create a new request for fallback auth.
  createRequest(clientRequest.url, clientRequest.callbackRESTApiHandler, async (fallbackRequest) => {
    // Hand off to call using fallback auth.
    await callWebServer(fallbackRequest);
  });
}


/**
 * Creates a client request object with:
 * authOptions - Auth configuration parameters for SSO.
 * authSSO - true if using SSO, otherwise false.
 * accessToken - The access token to the middle-tier server.
 * url - The URL of the REST API to call on the middle-tier server.
 * callbackRESTApiHandler - The function to pass the results of the REST API call.
 * callbackFunction - the function to pass the client request to when ready.
 *
 * Note that when the client request is created it will be passed to the callbackFunction. This is used because
 * we may need to pop up a dialog to sign in the user, which uses a callback approach.
 *
 * @param {*} callbackFunction The function to pass the client request to when ready.
 */
async function createRequest(url, restApiCallback, callbackFunction) {
  const clientRequest = {
    authOptions: {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    },
    authSSO: authSSO,
    accessToken: null,
    url: url,
    callbackRESTApiHandler: restApiCallback,
    callbackFunction: callbackFunction,
  };

  // Get access token.
  if (authSSO) {
    try {
      // Get access token from Office SSO.
      clientRequest.accessToken = await getAccessTokenFromSSO(clientRequest.authOptions);
      callbackFunction(clientRequest);
    } catch {
      // use fallback auth if SSO failed to get access token.
      switchToFallbackAuth(clientRequest);
    }
  } else {
    // Use fallback auth to get access token. 
    dialogFallback(clientRequest);
  }
}

/**
 * Returns the access token for using SSO auth. Throws an error if SSO fails.
 * @param {*} authOptions The configuration options for SSO.
 * @returns An access token to the middle-tier server for the signed in user.
 */
async function getAccessTokenFromSSO(authOptions) {
  if (authOptions === undefined) throw Error("authOptions parameter missing.");

  try {
    // The access token returned from getAccessToken only has permissions to your middle-tier server APIs,
    // and it contains the identity claims of the signed-in user.
    
    const accessToken = await Office.auth.getAccessToken(authOptions);
    return accessToken;
  } catch (error) {
    let fallbackRequired = handleSSOErrors(error);
    if (fallbackRequired) throw error; // Rethrow the error and caller will switch to fallback auth.
    return null; // Returning a null token indicates no need for fallback (an explanation about the error condition was shown by handleSSOErrors).
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

/**
 * Handles any error returned from the middle-tier server.
 * @param {*} err The error to process.
 * @returns {boolean} true if the caller should refresh the access token; otherwise false.
 */
function handleWebServerErrors(err) {
  let returnValue = false;
  // Our middle-tier server returns a type to help handle the known cases.
  switch (err.responseJSON.type) {
    case "Microsoft Graph":
      // An error occurred when the middle-tier server called Microsoft Graph.
      showMessage(
        "Error from Microsoft Graph: " +
        JSON.stringify(err.responseJSON.errorDetails)
      );
      returnValue = false;
      break;
    case "AADSTS500133": // expired token
      // On rare occasions the access token could expire after it was sent to the middle-tier server.
      // Microsoft identity platform will respond with
      // "The provided value for the 'assertion' is not valid. The assertion has expired."
      // Return true to indicate to caller they should refresh the token.
      returnValue = true;
      break;
    default:
      showMessage(
        "Unknown error from web server: " +
        JSON.stringify(err.responseJSON.errorDetails)
      );
      returnValue = false;
  }
  return false;
}
