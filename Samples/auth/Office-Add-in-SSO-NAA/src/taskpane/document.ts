// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* 
    This file provides the functionality to write data to the Office document. 
*/

/* global Excel Office PowerPoint Word */

//module.exports = writeFileNamesToOfficeDocument;
export { writeFileNamesToOfficeDocument };
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
        throw "Unsupported Office host application: This add-in only runs on Excel, PowerPoint, or Word.";
    }
  } catch (error) {
    throw Error(`Unable to add filenames to document. ${error}`);
  }
}

async function writeFileNamesToWorksheet(fileNames: string[]) {
  return Excel.run(async function (context) {
    const valuesToSet: string[][] = [];

    fileNames.forEach((fileName) => {
      valuesToSet.push([fileName]);
    });

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(`B5:B${5 + valuesToSet.length - 1}`);
    range.values = valuesToSet;
    range.format.autofitColumns();
    await context.sync();
  });
}

async function writeFileNamesToDocument(fileNames: string[]) {
  return Word.run(async (context) => {
    fileNames.forEach((name) => {
      context.document.body.insertParagraph(name, "End");
    });
    await context.sync();
  });
}

async function writeFileNamesToPresentation(fileNames: string[]) {
  const text = fileNames.join("\n");

  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.addTextBox(text, { width: 300, height: 300 });
    await context.sync();
  });
}
