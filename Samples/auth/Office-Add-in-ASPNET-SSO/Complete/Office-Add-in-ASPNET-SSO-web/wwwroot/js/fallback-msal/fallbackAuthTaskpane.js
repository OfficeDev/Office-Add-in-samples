// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file shows how to open a dialog and process any results sent back to the task pane.

const msalInstance2 = new msal.PublicClientApplication({
    auth: {
        clientId: "57b34432-e305-4b6d-9921-420a4c696a09",
        redirectUri: "https://localhost:7283/Account/Authorize",
        authority: "https://login.microsoftonline.com/organizations/"
    }
})

var loginDialog;
let storedCallbackFunction = null;
let storedClientRequest = null;

function dialogFallback(clientRequest) {
    storedClientRequest = clientRequest;
    var url = "/MicrosoftIdentity/Account/SignIn";
	showLoginPopup(url);
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
function processMessage(arg) {

    console.log("Message received in processMessage");
    let messageFromDialog = JSON.parse(arg.message);

        if (messageFromDialog.status === 'success') { 
            // We now have a valid SPA auth code.
            loginDialog.close();
            // Exchange the SPA auth code for an access token.
      //      const accessScope = "api://" + window.location.host + "/" + msalInstance2.clientId + "/access_as_user";
        //    const scopes = [accessScope];

          //  const tokenResponse = await getTokenFromCache(scopes);

            //storedClientRequest.accessToken = tokenResponse;
            storedClientRequest.accessToken = messageFromDialog.accessToken;
            storedClientRequest.callbackFunction(storedClientRequest);            
        }
        else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
            showMessage(JSON.stringify(error.toString()));
        }
}

///get Token
function getTokenPopup(spaCode) {

    var code = spaCode;
    const scopes = ["file.read"];

    console.log('MSAL: acquireTokenByCode hybrid parameters present');

    var authResult = msalInstance.acquireTokenByCode({
        code,
        scopes
    })
    console.log(authResult);

    return authResult

}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
	var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + url;

	// height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
	Office.context.ui.displayDialogAsync(fullUrl,
		{ height: 60, width: 30 }, function (result) {
			console.log("Dialog has initialized. Wiring up events");
			loginDialog = result.value;
			loginDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
		});
}
