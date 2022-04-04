/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to call APIs on your web server, which will then call Microsoft Graph.
 */

function makeWebServerApiCall(accessToken) {
    $.ajax({type: "GET", 
        url: "/getuserfilenames",
        headers: {"access_token": accessToken },
        cache: false
    }).done(function (response) {

        writeFileNamesToOfficeDocument(response)
        .then(function () { 
            showMessage("Your data has been added to the document."); 
        })
        .catch(function (error) {
            // The error from writeFileNamesToOfficeDocument will begin 
            // "Unable to add filenames to document."
            showMessage(error);
        });
    })
    .fail(function (errorResult) {
        // This error is relayed from `app.get('/getuserfilenames` in app.js file.
        showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
	});
}
