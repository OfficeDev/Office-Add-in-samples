/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.
 *
 * This file shows how to use the SSO API to get a bootstrap token.
 */

// If the add-in is running in Internet Explorer, the code must add support
// for Promises.
if (!window.Promise) {
  window.Promise = Office.Promise;
}

Office.onReady(function (info) {
  $(function () {
    $("#getGraphDataButton").on("click", getFileNameList);
  });
});

let retryGetAccessToken = 0;

async function getFileNameList(){
  try {
    
    // The access token returned from getAccessToken only has permissions to your web server APIs,
    // and the identity claims of the signed in user.
    let accessToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });

    let response = await callFileNameListAPI(accessToken);
    if (response!=null) writeFileNamesToOfficeDocument(response)
    .then(function () { 
        showMessage("Your data has been added to the document."); 
    })
    .catch(function (error) {
        // The error from writeFileNamesToOfficeDocument will begin 
        // "Unable to add filenames to document."
        showMessage(error);
    });
  } catch (exception) {
    // The only exceptions caught here are exceptions in your code in the try block
    // and errors returned from the call of `getAccessToken` above.
    if (exception.code) {
      handleClientSideErrors(exception);
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
  }
}

async function callFileNameListAPI(accessToken){
  return await $.ajax({type: "GET", 
  url: "/getuserfilenames",
  headers: { Authorization: "Bearer " + accessToken },
  cache: false
}).done(function (response) {
return response;
})
.fail(function (errorResult) {
  // This error is relayed from `app.get('/getuserfilenames` in app.js file.
  showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
  return null;
});
}
function handleClientSideErrors(error) {
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

// function handleAADErrors(exchangeResponse) {
//   // On rare occasions the bootstrap token is unexpired when Office validates it,
//   // but expires by the time it is sent to AAD for exchange. AAD will respond
//   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
//   // Retry the call of getAccessToken (no more than once). This time Office will return a
//   // new unexpired bootstrap token.
//   if (
//     exchangeResponse.error_description.indexOf("AADSTS500133") !== -1 &&
//     retryGetAccessToken <= 0
//   ) {
//     retryGetAccessToken++;
//     getGraphData();
//   } else {
//     // For all other AAD errors, fallback to non-SSO sign-in.
//     // For debugging:
//     // showMessage("AAD ERROR: " + JSON.stringify(exchangeResponse));
//     dialogFallback();
//   }
// }
