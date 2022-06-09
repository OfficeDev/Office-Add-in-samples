/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.
 *
 */

// global to track if we are using SSO or the fallback auth.
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
 * Requests a call to the web server /getuserfilenames that
 * gets up to 10 file names listed in the user's OneDrive.
 * When the call is completed, it will call the clientRequest.callbackHandler.
 */
function getFileNameList() {
  createRequest((clientRequest) => {
    clientRequest.url = "/getuserfilenames";
    clientRequest.callbackHandler = handleGetFileNameResponse;
    callWebServer(clientRequest);
  });
}

/**
 * Handler for the returned response from the server API call to get file names.
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
  } else showMessage("A null response was returned to handleGetFileNameReponse.");
}

/**
 * Checks to see if SSO or fallback auth is used.
 * Then calls the correct method to pass the client request to the server.
 * 
 * @param {*} clientRequest Contains information for calling an API on the server.
 */
function callWebServer(clientRequest) {
  if (clientRequest.authSSO) {
    callServerWithSSO(clientRequest);
  } else {
    callServerWithFallback(clientRequest);
  }
}

/**
 * Calls the sever API using SSO as the auth approach.
 * If the call is successful, the clientRequest.callbackHandler method is called to handle the results.
 * 
 * @param {*} clientRequest Contains information for calling an API on the server.
 */
async function callServerWithSSO(clientRequest) {
  var result = null;
  var tokenNeeded = true; // The following loop will call getAccessToken again in some error scenarios. retryGetAccessToken is configure to only retry the loop once.
  var retryGetAccessToken = 0; // Use when getAccessToken is called repeatedly to control recursion depth.

  while (tokenNeeded && retryGetAccessToken <= 1) {
    // Call our web server with requested url
    try {
      await $.ajax({
        type: "GET",
        url: clientRequest.url,
        headers: { Authorization: "Bearer " + clientRequest.accessToken },
        cache: false,
        success: function (data) {
          tokenNeeded = false;
          result = data;
          // call the handler method 
          clientRequest.callbackHandler(result);
        },
      });
    } catch (error) {
      // We only handle errors returned by our web server (500).
      if (error.statusText === "Internal Server Error") {
        tokenNeeded = handleWebServerErrors(error);
      } else console.log(JSON.stringify(error)); // Log anything else.
    }
    retryGetAccessToken++;
  }
  if (tokenNeeded) {
    // We exceeded the loop count and could not obtain an SSO token.
    authSSO = false;
    // We need to create a new request
    createRequest((fallbackRequest) => {
      fallbackRequest.url = clientRequest.url;
      fallbackRequest.callbackHandler = clientRequest.callbackHandler;
      // Hand off to call using fallback auth.
      callServerWithFallback(fallbackRequest);
      return;
    });
  }
}

/**
 * Calls the sever API using the fallback auth approach.
 * @param {*} clientRequest Contains information for calling an API on the server.
 */
async function callServerWithFallback(clientRequest) {
  // Call our web server with requested url
  try {
    await $.ajax({
      type: "GET",
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
      success: function (data) {
        result = data;
        clientRequest.callbackHandler(result);
      },
    });
  } catch (error) {
    // We only handle errors returned by our web server (500).
    if (error.statusText === "Internal Server Error") {
      handleWebServerErrors(error);
    } else console.log(JSON.stringify(error)); // Log anything else.
  }
}

/**
 * Creates a client request object with:
 * authOptions - Auth configuration parameters for SSO.
 * authSSO - true if using SSO, otherwise false.
 * accessToken - The access token to the server.
 * url - The URL of the REST API to call on the server.
 * callbackHandler - The function to pass the results of the REST API call.
 * callbackFunction - the function to pass the client request to when ready.
 * 
 * Note that when the client request is created it will be passed to the callbackFunction. This is used because
 * we may need to pop up a dialog to sign in the user, which uses a callback approach.
 * 
 * @param {*} callbackFunction The function to pass the client request to when ready. 
 */
async function createRequest(callbackFunction) {
  const clientRequest = {
    authOptions: {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    },
    authSSO: authSSO,
    accessToken: null,
    url: null,
    callbackHandler: null,
    callbackFunction: callbackFunction
  }

  if (authSSO) {
    // Use SSO approach.
    try {
      clientRequest.accessToken = await getAccessTokenFromSSO();
      callbackFunction(clientRequest);
    } catch {
      // use fallback auth if SSO failed.
      authSSO = false;
      dialogFallback(clientRequest);
    }
  } else {
    // Use fallback auth approach
    dialogFallback(clientRequest);
  }
}

/**
 * Returns the access token for using SSO auth. Throws an error if SSO fails. 
 * @param {*} authOptions The configuration options for SSO.
 * @returns An access token to the server for the signed in user.
 */
async function getAccessTokenFromSSO(authOptions) {
  if (authOptions === undefined) {
    // Set up default auth options.
    authOptions = {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    };
  }
  let accessToken = null;

  try {
    // The access token returned from getAccessToken only has permissions to your web server APIs,
    // and it contains the identity claims of the signed-in user.
    let authParam = { ...authOptions };
    accessToken = await Office.auth.getAccessToken(authParam);
  } catch (error) {
    let fallbackRequired = handleSSOErrors(error);
    if (fallbackRequired) throw (error); // Rethrow the error and caller will switch to fallback auth.
  }
  return accessToken;
}


/**
 * Handles any error returned from getAccessToken.
 * @param {*} err The error to process.
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
 * Handles any error returned from the web server.
 * @param {*} err The error to process.
 * @returns {boolean} true if the caller should attempt to retry getAccessToken; otherwise false.
 */
function handleWebServerErrors(err) {
  // Our web server returns a type to help handle the known cases.
  switch (err.responseJSON.type) {
    case "Microsoft Graph":
      // An error occurred when the web server called Microsoft Graph.
      showMessage(
        "Error from Microsoft Graph: " +
        JSON.stringify(err.responseJSON.errorDetails)
      );
      break;
    case "AADSTS500133": // expired token
      // On rare occasions the access token could expire after it was sent to the server.
      // Microsoft identity platform will respond with
      // "The provided value for the 'assertion' is not valid. The assertion has expired."
      // return parameters for the calling function to retry the callWebServerAPI method recursively to try to get a refreshed SSO token.
      return true; // Indicate to retry call to getAccessToken.
      break;
    default:
      showMessage(
        "Unknown error from web server: " +
        JSON.stringify(err.responseJSON.errorDetails)
      );
      // We should use fallback auth from this point if an unknown error occurred while using the SSO token.
      if (authSSO) {
        authSSO = false;
        dialogFallback();
      }
      return false;
  }
  return false; // Indicate no need to retry call to getAccessToken.
}