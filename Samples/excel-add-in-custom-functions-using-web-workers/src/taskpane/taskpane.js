/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Excel, Office */

Office.onReady(() => {
  document.getElementById("runCFWithoutWebWorker").onclick = () => tryCatch(runCFWithoutWebWorker);
  document.getElementById("runCFWithWebWorker").onclick = () => tryCatch(runCFWithWebWorker);
});

async function runCFWithoutWebWorker() {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    const sheet = context.workbook.worksheets.add("Sample");

    let range = sheet.getRange("A1");
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

    let range = sheet.getRange("A1");
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

function showStatus(message, isError) {
  let status = document.getElementById("status");
  // Clear previous content.
  status.innerHTML = "";

  // Create the container div.
  let statusCard = document.createElement("div");
  statusCard.className = `status-card ms-depth-4 ${isError ? "error-msg" : "success-msg"}`;

  // Create and append the first paragraph.
  let p1 = document.createElement("p");
  p1.className = "ms-fontSize-24 ms-fontWeight-bold";
  p1.textContent = isError ? "An error occurred" : "";
  statusCard.appendChild(p1);

  // Create and append the second paragraph.
  let p2 = document.createElement("p");
  p2.className = "ms-fontSize-16 ms-fontWeight-regular";
  p2.textContent = message;
  statusCard.appendChild(p2);

  // Append the status card to the status element.
  status.appendChild(statusCard);
}
