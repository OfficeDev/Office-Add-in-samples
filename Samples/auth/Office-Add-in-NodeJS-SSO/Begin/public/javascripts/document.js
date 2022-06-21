// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.


/* 
    This file provides the functionality to write data to the Office document. 
*/

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
        var sheet = context.workbook.worksheets.getActiveWorksheet();

        var filenames = [];
        var i;
        for (i = 0; i < result.length; i++) {
            var innerArray = [];
            innerArray.push(result[i]);
            filenames.push(innerArray);
        }

        var rangeAddress = "B5:B" + (5 + (result.length - 1)).toString();
        var range = sheet.getRange(rangeAddress);
        range.values = filenames;
        range.format.autofitColumns();

        return context.sync();
    });
}

function writeFileNamesToDocument(result) {

     return Word.run(function (context) {

        var documentBody = context.document.body;
        for (var i = 0; i < result.length; i++) {
            documentBody.insertParagraph(result[i], "End");
        }

        return context.sync();
    });
}

function writeFileNamesToPresentation(result) {

    var fileNames = "";
    for (var i = 0; i < result.length; i++) {
        fileNames += result[i] + '\n';
    }

    Office.context.document.setSelectedDataAsync(
        fileNames,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                throw asyncResult.error.message;
            }
        }
    );
}