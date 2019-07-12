// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

"use strict";

let dialog;

Office.initialize = function () {
    $(document).ready(function () {
        app.initialize();

        $("#getOneDriveFilesButton").click(getFileNamesFromGraph);
        $("#logoutO365PopupButton").click(logout);        
    });
};

function getFileNamesFromGraph() {

    $("#instructionsContainer").hide();
    $("#waitContainer").show();

    $.ajax({
        url: "/files/onedrivefiles",
        type: "GET"
    })
    .done(function (result) {
        writeFileNamesToMessage(result)
            .then(function () {
                $("#waitContainer").hide();
                $("#finishedContainer").show();
            })
            .catch(function (error) {
                app.showNotification(error.toString());
            });
    })
        .fail(function (result) {
            app.showNotification("Cannot get data from MS Graph: " + result.toString());
    });
}

function writeFileNamesToMessage(graphData) {

    // Office.Promise is an alias of OfficeExtension.Promise. Only the alias
    // can be used in an Outlook add-in.
    return new Office.Promise(function (resolve, reject) {
        try {
            Office.context.mailbox.item.body.getTypeAsync(
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        app.showNotification(result.error.message);
                    }
                    else {
                        // Successfully got the type of item body.
                        if (result.value === Office.MailboxEnums.BodyType.Html) {

                            // Body is of type HTML.
                            var htmlContent = createHtmlContent(graphData);

                            Office.context.mailbox.item.body.setSelectedDataAsync(
                                htmlContent, { coercionType: Office.CoercionType.Html },
                                function (asyncResult) {
                                    if (asyncResult.status ===
                                        Office.AsyncResultStatus.Failed) {
                                        console.log(asyncResult.error.message);
                                    }
                                    else {
                                        console.log("Successfully set HTML data in item body.");
                                    }
                                });
                        }
                        else {
                            // Body is of type text. 
                            var textContent = createTextContent(graphData);

                            Office.context.mailbox.item.body.setSelectedDataAsync(
                                textContent, { coercionType: Office.CoercionType.Text },
                                function (asyncResult) {
                                    if (asyncResult.status ===
                                        Office.AsyncResultStatus.Failed) {
                                        console.log(asyncResult.error.message);
                                    }
                                    else {
                                        console.log("Successfully set text data in item body.");
                                    }
                                });
                        }
                    }
                });
            resolve();
        }
        catch (error) {
            reject(Error("Unable to add filenames to document. " + error));
        }
    });
}

function createHtmlContent(data) {

    var bodyContent = "<html><head></head><body>";

    for (var i = 0; i < data.length; i++) {
        bodyContent += "<p>" + data[i] + "</p>";
    }
    bodyContent += "</body></html >";

    return bodyContent;
}

function createTextContent(data) {

    var bodyContent = "";
    for (var i = 0; i < data.length; i++) {
        bodyContent += data[i] + "\n";
    }

    return bodyContent;
}

function logout() {

    Office.context.ui.displayDialogAsync('https://localhost:44301/azureadauth/logout',
        { height: 30, width: 30 }, function (result) {           
            dialog = result.value;
            dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processLogoutMessage);
        });
}

function processLogoutMessage(messageFromLogoutDialog) {

    if (messageFromLogoutDialog.message === "success") {
        dialog.close();
        document.location.href = "/home/index";
    }
    else {
        dialog.close();
        app.showNotification("Not able to logout: " + messageFromLogoutDialog.toString());
    }
}