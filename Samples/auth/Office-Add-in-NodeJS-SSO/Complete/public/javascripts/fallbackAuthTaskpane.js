/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to open a dialog and process any results sent back to the task pane.
 */

var loginDialog;
let storedCallbackFunction = null;

function dialogFallback(callbackFunction) {
    storedCallbackFunction = callbackFunction;
    var url = "/dialog.html"; 
	showLoginPopup(url);
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
function processMessage(arg) {

    console.log("Message received in processMessage: " + JSON.stringify(arg));
    let messageFromDialog = JSON.parse(arg.message);

        if (messageFromDialog.status === 'success') { 
            // We now have a valid access token.
            loginDialog.close();
            const url = localStorage.getItem("url");
            storedCallbackFunction(messageFromDialog.result);
            //makeGraphApiCall(messageFromDialog.result);
        }
        else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
            showMessage(JSON.stringify(error.toString()));
        }
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
