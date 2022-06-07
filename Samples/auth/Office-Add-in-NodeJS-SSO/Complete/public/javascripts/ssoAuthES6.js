/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.
 *
 */

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
 * The file names are inserted into the document.
 */
function getFileNameList() {
  getCorrectAccessToken((accessToken) => {
    callWebServerAPI2("/getuserfilenames", accessToken).then((response) => {
      if (response != null)
        writeFileNamesToOfficeDocument(response)
          .then(function () {
            showMessage("Your OneDrive filenames are added to the document.");
          })
          .catch(function (error) {
            // The error from writeFileNamesToOfficeDocument will begin
            // "Unable to add filenames to document."
            showMessage(error);
          });
    });
  });
}

let authSSO = false;

// When using the fallback approach, the fallback dialog will call this function again with the access token.
async function callWebServerAPI2(url, accessToken) {
  // Call our web server with requested url
  try {
    await $.ajax({
      type: "GET",
      url: url,
      headers: { Authorization: "Bearer " + accessToken },
      cache: false,
      success: function (data) {
        tokenNeeded = false;
        result = data;
      },
    });
  } catch (error) {
    // We only handle errors returned by our web server (500).
    if (error.statusText === "Internal Server Error") {
      tokenNeeded = handleWebServerErrors(error);
    } else console.log(JSON.stringify(error)); // Log anything else.
  }
  return result;
}

// get token based on if we are using SSO or fallback.
function getCorrectAccessToken(callbackFunction) {
  if (authSSO) {
    //return the sso token
  } else {
    //return the fallback token
    dialogFallback(callbackFunction);
  }
}

/**
 * Calls the add-in's web server API specified by the url. Only makes GET calls.
 *
 * @param {*} url The url specifying the REST API name to call.
 * @returns The response from the server.
 */
async function callWebServerAPI(url, authOptions) {
  dialogFallback();
  if (authOptions === undefined) {
    // Set up default auth options.
    authOptions = {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    };
  }
  var accessToken = null;
  var result = null;
  var tokenNeeded = true; // The following loop will call getAccessToken again in some error scenarios. retryGetAccessToken is configure to only retry the loop once.
  var retryGetAccessToken = 0; // Use when getAccessToken is called repeatedly to control recursion depth.

  while (tokenNeeded && retryGetAccessToken <= 1) {
    try {
      // The access token returned from getAccessToken only has permissions to your web server APIs,
      // and it contains the identity claims of the signed-in user.
      let authParam = { ...authOptions };
      accessToken = await Office.auth.getAccessToken(authParam);
    } catch (error) {
      handleSSOErrors(error);
    }

    // Call our web server with requested url
    try {
      await $.ajax({
        type: "GET",
        url: url,
        headers: { Authorization: "Bearer " + accessToken },
        cache: false,
        success: function (data) {
          tokenNeeded = false;
          result = data;
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
    dialogFallback();
  }
  return result;
}

/**
 * Handles any error returned from getAccessToken.
 * @param {*} err The error to process
 */
function handleSSOErrors(err) {
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
      dialogFallback();
      break;
  }
}

/**
 * Handles any error returned from the web server.
 * @param {*} err The error to process
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
      dialogFallback();
      return false;
  }
  return false; // Indicate no need to retry call to getAccessToken.
}