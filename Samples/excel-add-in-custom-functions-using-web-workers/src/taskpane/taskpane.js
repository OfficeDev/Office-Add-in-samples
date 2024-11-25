/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(() => {
  document.getElementById("runCFWithoutWebWorker").onclick = () => tryCatch(runCFWithoutWebWorker);
  document.getElementById("runCFWithWebWorker").onclick = () => tryCatch(runCFWithWebWorker);
});


async function runCFWithoutWebWorker() {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    const sheet = context.workbook.worksheets.add("Sample");

    let range = sheet.getRange('A1');
    range.values = [["=WebWorkerSample.TEST_UI_THREAD(20000)"]];
    range.calculate();

    sheet.activate();
    await context.sync();
  });
}

async function runCFWithWebWorker() {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    const sheet = context.workbook.worksheets.add("Sample");

    let range = sheet.getRange('A1');
    range.values = [["=WebWorkerSample.TEST(20000)"]];
    range.calculate();

    sheet.activate();
    await context.sync();
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    showStatus(error, true);
  }
}
