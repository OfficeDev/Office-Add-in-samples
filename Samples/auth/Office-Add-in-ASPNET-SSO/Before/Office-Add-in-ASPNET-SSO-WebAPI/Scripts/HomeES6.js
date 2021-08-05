// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides functions to get ask the Office host to get an access token to the add-in
	and to pass that token to the server to get Microsoft Graph data. 
*/

// To support IE (which uses ES5), we have to create a Promise object for the global context.
if (!window.Promise) {
    window.Promise = Office.Promise;
}

Office.initialize = function (reason) {

    $(document).ready(function () {
        $('#getGraphDataButton').click(getGraphData);
    });
};




// Displays the data, assumed to be an array.
function showResult(data) {

	// Use jQuery text() method which automatically encodes values that are passed to it,
    // in order to protect against injection attacks.
	$.each(data, function (i) {
		var li = $('<li/>').addClass('ms-ListItem').appendTo($('#file-list'));
		var outerSpan = $('<span/>').addClass('ms-ListItem-secondaryText').appendTo(li);
		$('<span/>').addClass('ms-fontColor-themePrimary').appendTo(outerSpan).text(data[i]);
	});
}

function logError(result) {
	console.log("Status: " + result.status);
	console.log("Code: " + result.error.code);
	console.log("Name: " + result.error.name);
	console.log("Message: " + result.error.message);
}

// Dialog API

var loginDialog;
var redirectTo = "/files/index";

function dialogFallback() {

    var url = "/azureadauth/login"; 
	showLoginPopup(url);
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
function processMessage(arg) {

    console.log("Message received in processMessage: " + JSON.stringify(arg));
    let message = JSON.parse(arg.message);

    if (message.status === "success") {
        // We now have a valid access token.
        loginDialog.close();
        getData("/api/files", message.accessToken);
    } else {
        // Something went wrong with authentication or the authorization of the web application.
        loginDialog.close();
        showResult(["Unable to successfully authenticate user or authorize application. Error is: " + message.error]);
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