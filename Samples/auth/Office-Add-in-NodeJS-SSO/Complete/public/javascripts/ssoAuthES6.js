/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.
 *
 */

let retryGetAccessToken = 0; // Use when getAccessToken is call repeatedly to control recursion depth

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
async function getFileNameList() {
  const response = await callWebServerAPI("/getuserfilenames");
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
}

/**
 * Calls the add-in's web server API specified by the url. Only makes GET calls.
 *
 * @param {*} url The url specifying the REST API name to call.
 * @returns The response from the server.
 */
async function callWebServerAPI(url, authOptions) {
  if (authOptions === undefined) {
    // Set up default auth options.
    let authOptions = {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    };
  }
  let accessToken = null;

  // Get the SSO access token from Office
  try {
    // The access token returned from getAccessToken only has permissions to your web server APIs,
    // and it contains the identity claims of the signed-in user.
    let accessToken = await Office.auth.getAccessToken(authOptions);
    retryGetAccessToken = 0; // If success we can reset this counter
  } catch (error) {
    handleSSOErrors(error);
  }

  // Call our web server using the SSO access token
  try {
    const response = await $.ajax({
      type: "GET",
      url: url,
      headers: { Authorization: "Bearer " + accessToken },
      cache: false,
    });
    return response;
  } catch (e) {
    handleWebServerErrors(e);
  }
}

/**
 * Handles any error returned from getAccessToken.
 * @param {*} error The error to process
 */
function handleSSOErrors(error) {
  switch (error.code) {
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

function handleWebServerErrors(e) {
  // Our special handling on the server will cause the result that is returned
  // from a AADSTS50076 (a 2FA challenge) to have a Message property but no ExceptionMessage.
  var message = e.responseJSON.Message;

  // Results from other errors (other than AADSTS50076) will have an ExceptionMessage property.
  var exceptionMessage = result.responseJSON.ExceptionMessage;

  if (
    exceptionMessage &&
    e.Message.indexOf("AADSTS500133") !== -1 &&
    retryGetAccessToken <= 0
  ) {
    // On rare occasions the access token could expire after it was sent to the server.
    // Microsoft identity platform will respond with
    // "The provided value for the 'assertion' is not valid. The assertion has expired."
    // Call this method recursively to try to get an SSO token again.
    callWebServerAPI(url, authOptions);
  } else if (message) {
    // Microsoft Graph requires an additional form of authentication. Have the Office host
    // get a new token using the Claims string, which tells Microsoft identity platform to
    // prompt the user for all required forms of authentication.
    if (message.indexOf("AADSTS50076") !== -1 && retryGetAccessToken <= 0) {
      const claims = JSON.parse(message).Claims;
      const claimsAsString = JSON.stringify(claims);
      authOptions.authChallenge = claimsAsString;
      callWebServerAPI(url, {
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
    }
  } else {
    // For debugging:
    // showResult(["Microsoft identity platform error: " + JSON.stringify(exceptionMessage)]);

    // For all other Microsoft identity platform errors, fallback to non-SSO sign-in.
    dialogFallback();
  }
}
