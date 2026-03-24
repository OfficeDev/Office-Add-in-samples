/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Excel, Office */

Office.onReady(() => {
  document.getElementById("setup").onclick = () => tryCatch(setup);
});

async function setup() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Set up sample data in A1:A3.
    sheet.getRange("A1").values = [["Value"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A2").values = [[42]];
    sheet.getRange("A3").values = [[100]];

    // Use the synchronous custom function in B2:B3 to read cells A2 and A3.
    sheet.getRange("B1").values = [["GETCELLVALUE Result"]];
    sheet.getRange("B1").format.font.bold = true;
    sheet.getRange("B2").values = [['=SyncCFSample.GETCELLVALUE("A2")']];
    sheet.getRange("B3").values = [['=SyncCFSample.GETCELLVALUE("A3")']];

    sheet.getUsedRange().format.autofitColumns();
    await context.sync();

    document.getElementById("status").textContent =
      "Data ready. Cells B2:B3 use GETCELLVALUE to synchronously read cells A2:A3.";
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
