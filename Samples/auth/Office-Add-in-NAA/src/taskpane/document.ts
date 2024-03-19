// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* 
    This file provides the functionality to write data to the Office document. 
*/

//module.exports = writeFileNamesToOfficeDocument;
export { writeFileNamesToOfficeDocument }
function writeFileNamesToOfficeDocument(fileNameList: string[]) {
  try {
    switch (Office.context.host) {
      case Office.HostType.Excel:
        return writeFileNamesToWorksheet(fileNameList);
      case Office.HostType.Word:
        return writeFileNamesToDocument(fileNameList);
      case Office.HostType.PowerPoint:
        return writeFileNamesToPresentation(fileNameList);
      default:
        throw 'Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.';
    }
  } catch (error) {
    throw Error('Unable to add filenames to document. ' + error.toString());
  }
}

async function writeFileNamesToWorksheet(fileNameList: string[]) {
  return Excel.run(async function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Build correct array structure to update the range later.
    const fileNames = [];
    for (let i = 0; i < fileNameList.length; i++) {
      var innerArray = [];
      innerArray.push(fileNameList[i]);
      fileNames.push(innerArray);
    }

    // Update the range.
    const rangeAddress = 'B5:B' + (5 + (fileNameList.length - 1)).toString();
    const range = sheet.getRange(rangeAddress);
    range.values = fileNames;
    range.format.autofitColumns();

    await context.sync();
  });
}

async function writeFileNamesToDocument(fileNameList: string[]) {
  return Word.run(function (context) {
    const documentBody = context.document.body;
    for (let i = 0; i < fileNameList.length; i++) {
      documentBody.insertParagraph(fileNameList[i], 'End');
    }

    return context.sync();
  });
}

async function writeFileNamesToPresentation(fileNameList: string[]) {
  let fileNames = '';
  for (var i = 0; i < fileNameList.length; i++) {
    fileNames += fileNameList[i] + '\n';
  }

  Office.context.document.setSelectedDataAsync(
    fileNames,
    function (asyncfileNameList) {
      if (asyncfileNameList.status === Office.AsyncResultStatus.Failed) {
        throw asyncfileNameList.error.message;
      }
    }
  ); 
}
