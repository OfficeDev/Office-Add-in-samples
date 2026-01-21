// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* 
    This file provides the functionality to write data to the Office document. 
*/

function writeFileNamesToOfficeDocument(result) {
  try {
    switch (Office.context.host) {
      case 'Excel':
        return writeFileNamesToWorksheet(result);
      case 'Word':
        return writeFileNamesToDocument(result);
      case 'PowerPoint':
        return writeFileNamesToPresentation(result);
      default:
        throw 'Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.';
    }
  } catch (error) {
    throw Error('Unable to add filenames to document. ' + error.toString());
  }
}

async function writeFileNamesToWorksheet(result) {
  return Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    var filenames = [];
    var i;
    for (i = 0; i < result.length; i++) {
      var innerArray = [];
      innerArray.push(result[i]);
      filenames.push(innerArray);
    }

    var rangeAddress = 'B5:B' + (5 + (result.length - 1)).toString();
    var range = sheet.getRange(rangeAddress);
    range.values = filenames;
    range.format.autofitColumns();

    return context.sync();
  });
}

async function writeFileNamesToDocument(result) {
  return Word.run(function (context) {
    var documentBody = context.document.body;
    for (var i = 0; i < result.length; i++) {
      documentBody.insertParagraph(result[i], 'End');
    }

    return context.sync();
  });
}

async function writeFileNamesToPresentation(result) {
  var fileNames = '';
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
