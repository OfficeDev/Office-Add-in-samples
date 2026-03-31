// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

"use strict";

async function writeFileNamesToOfficeDocument(result) {
        try {
            switch (Office.context.host) {
                case "Excel":
                    await writeFileNamesToWorksheet(result);
                    break;
                case "Word":
                    await writeFileNamesToDocument(result);
                    break;
                case "PowerPoint":
                    await writeFileNamesToPresentation(result);
                    break;
                default:
                    throw "Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.";
            }
        }
        catch (error) {
            throw new Error("Unable to add filenames to document. " + error.toString());
        }
}

async function writeFileNamesToWorksheet(result) {
     await Excel.run( (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const filenames = [];
        
        for (let i = 0; i < result.length; i++) {
            let innerArray = [];
            innerArray.push(result[i].name);
            filenames.push(innerArray);
        }

        const rangeAddress = "B5:B" + (5 + (result.length - 1)).toString();
        const range = sheet.getRange(rangeAddress);
        range.values = filenames;
        range.format.autofitColumns();

        return context.sync();
    });
}

async function writeFileNamesToDocument(result) {

     await Word.run( (context) => {

        const documentBody = context.document.body;
        for (let i = 0; i < result.length; i++) {
            documentBody.insertParagraph(result[i].name, "End");
        }

        return context.sync();
    });
}

function writeFileNamesToPresentation(result) {
    return new Promise(function (resolve, reject) {
        let fileNames = "";
        for (let i = 0; i < result.length; i++) {
            fileNames += result[i].name + '\n';
        }

        Office.context.document.setSelectedDataAsync(
            fileNames,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    reject(asyncResult.error.message);
                }
            }
        );
    });
}

