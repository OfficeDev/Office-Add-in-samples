/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Excel, Office */

Office.onReady(() => {
  document.getElementById("run").onclick = () => tryCatch(run);
});

async function run() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // A1:A2 — label and target threshold.
    sheet.getRange("A1").values = [["Target"]];
    sheet.getRange("A2").values = [[70]];
    sheet.getRange("A1:A2").format.font.bold = true;

    // B1:B8 — header and sample scores.
    sheet.getRange("B1").values = [["Score"]];
    sheet.getRange("B1").format.font.bold = true;
    sheet.getRange("B2:B8").values = [[45], [82], [91], [60], [78], [55], [88]];

    // Apply a conditional format rule on B2:B8.
    // The formula uses the synchronous custom function GETCELLVALUE to read
    // the target from A2 during calculation. Scores exceeding the target
    // are highlighted.
    const range = sheet.getRange("B2:B8");
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    cf.custom.rule.formula = '=B2>SyncCFConditionalFormat.GETCELLVALUE("$A$2")';
    cf.custom.format.fill.color = "#FFC7CE";
    cf.custom.format.font.color = "#9C0006";

    sheet.getUsedRange().format.autofitColumns();
    await context.sync();

    document.getElementById("status").textContent =
      "Done. Scores in B2:B8 that exceed the target in A2 are highlighted. Change A2 to see the conditional format update live.";
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
