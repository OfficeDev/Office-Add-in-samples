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
        writeFileNamesToOfficeDocument(result)
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

function writeFileNamesToOfficeDocument(result) {

    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            switch (Office.context.host) {
                case "Excel":
                    writeFileNamesToWorksheet(result);
                    break;
                case "Word":
                    writeFileNamesToDocument(result);
                    break;
                case "PowerPoint":
                    writeFileNamesToPresentation(result);
                    break;
                default:
                    throw "Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.";
            }
            resolve();
        }
        catch (error) {
            reject(Error("Unable to add filenames to document. " + error.toString()));
        }
    });    
}

function writeFileNamesToWorksheet(result) {
    
     return Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        const data = [
             [result[0]],
             [result[1]],
             [result[2]]];

        const range = sheet.getRange("B5:B7");
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
    });
}

function writeFileNamesToDocument(result) {

     return Word.run(function (context) {

        const documentBody = context.document.body;
        for (let i = 0; i < result.length; i++) {
            documentBody.insertParagraph(result[i], "End");
        }

        return context.sync();
    });
}

function writeFileNamesToPresentation(result) {

    const fileNames = result[0] + '\n' + result[1] + '\n' + result[2];

    Office.context.document.setSelectedDataAsync(
        fileNames,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                throw asyncResult.error.message;
            }
        }
    );
}

function logout() {

    Office.context.ui.displayDialogAsync('https://localhost:44301/azureadauth/logout',
        { height: 30, width: 30 }, function (result) {           
            dialog = result.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processLogoutMessage);
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