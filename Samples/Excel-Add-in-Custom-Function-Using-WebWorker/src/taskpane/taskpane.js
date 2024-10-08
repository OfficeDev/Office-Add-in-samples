/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  // Check that we loaded into Excel
  if (info.host === Office.HostType.Excel) {
    //let userClientId = 'YOUR_APP_ID_HERE'; //Register your app at https://aad.portal.azure.com/
    //localStorage.setItem('client-id', userClientId);

    //document.getElementById("sendEmail").onclick = checkClientID;
    document.getElementById("executeCFWithoutWebWorker").onclick = executeCFWithoutWebWorker;
    document.getElementById("executeCFWithWebWorker").onclick = executeCFWithWebWorker;
  }
});


// Execute Custom Function without WebWorker
async function executeCFWithoutWebWorker() {
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      const sheet = context.workbook.worksheets.add("Sample");

      let range = sheet.getRange('A1');
      range.values = [["=WebWorkerSample.TEST_UI_THREAD(20000)"]];
      range.calculate();

      sheet.activate();
      await context.sync();
    });
  } catch (error) {
    showStatus(`Exception when executeCFWithoutWebWorker: ${JSON.stringify(error)}`, true);
  }
}

// Execute Custom Function with WebWorker
async function executeCFWithWebWorker() {
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      const sheet = context.workbook.worksheets.add("Sample");

      let range = sheet.getRange('A1');
      range.values = [["=WebWorkerSample.TEST(20000)"]];
      range.calculate();

      sheet.activate();
      await context.sync();
    });
  } catch (error) {
    showStatus(`Exception when executeCFWithWebWorker: ${JSON.stringify(error)}`, true);
  }
}

