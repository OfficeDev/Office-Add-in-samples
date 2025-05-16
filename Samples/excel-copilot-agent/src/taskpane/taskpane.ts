/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(() => {
  document.getElementById("add-sample-data").onclick = () => tryCatch(addSampleData);
 });

async function addSampleData() {
  await Excel.run(async (context) => {
    showStatus("Running...", false);
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    context.workbook.worksheets.add("Sample");
    const sheet = context.workbook.worksheets.getItem("Sample");

    let expensesTable = sheet.tables.add("A1:H1", true);
    expensesTable.name = "SalesTable";
    expensesTable.getHeaderRowRange().values = [
      [
        "Date",
        "Location",
        "Temperature",
        "Leaflets",
        "Price",
        "Lemon Drink Sales",
        "Orange Drink Sales",
        "Total Sales",
      ],
    ];
    expensesTable.rows.add(null, [
      ["7/1/2016", "Park", 70, 90, 0.25, 9409, 4489, 0],
      ["7/2/2016", "Park", 72, 90, 0.25, 9604, 4489, 0],
      ["7/3/2016", "Park", 71, 104, 0.25, 12100, 5929, 0],
      ["7/4/2016", "Beach", 76, 98, 0.25, 17956, 9801, 0],
      ["7/5/2016", "Beach", 78, 135, 0.25, 25281, 13924, 0],
      ["7/6/2016", "Beach", 82, 90, 0.25, 10609, 4761, 0],
      ["7/6/2016", "Beach", 82, 90, 0.25, 10609, 4761, 0],
      ["7/7/2016", "Beach", 81, 135, 0.25, 20449, 10201, 0],
      ["7/8/2016", "Beach", 82, 113, 0.25, 15129, 7396, 0],
      ["7/9/2016", "Beach", 80, 126, 0.25, 17956, 9025, 0],
      ["7/10/2016", "Beach", 82, 131, 0.25, 19600, 9604, 0],
      ["7/11/2016", "Beach", 83, 135, 0.25, 26244, 14400, 0],
      ["7/12/2016", "Beach", 84, 99, 0.25, 16900, 9025, 0],
      ["7/13/2016", "Beach", 77, 99, 0.25, 11881, 5625, 0],
      ["7/14/2016", "Beach", 78, 113, 0.25, 14884, 7225, 0],
      ["7/15/2016", "Beach", 75, 108, 0.5, 9604, 3844, 0],
      ["7/16/2016", "Beach", 74, 90, 0.5, 6561, 2500, 0],
      ["7/17/2016", "Beach", 77, 126, 0.5, 13225, 5776, 0],
      ["7/18/2016", "Park", 81, 122, 0.5, 17161, 8464, 0],
      ["7/19/2016", "Park", 78, 113, 0.5, 14884, 7225, 0],
      ["7/20/2016", "Park", 70, 120, 0.5,5041, 1764, 0],
      ["7/21/2016", "Park", 77, 90, 0.5, 6889, 2500, 0],
      ["7/22/2016", "Park", 80, 108, 0.5, 12544, 5625, 0],
      ["7/23/2016", "Park", 81, 117, 0.5, 14400, 6724, 0],
      ["7/24/2016", "Park", 82, 117, 0.5, 14641, 6724, 0],
      ["7/25/2016", "Park", 84, 135, 0.5, 24336, 12769, 0],
      ["7/26/2016", "Park", 83, 158, 0.35, 30976, 16641, 0],
      ["7/27/2016", "Park", 80, 99, 0.35, 10816, 4624, 0],
      ["7/28/2016", "Park", 82, 90, 0.35, 9216, 3969, 0],
      ["7/29/2016", "Park", 81, 95, 0.35, 10000, 4356, 0],
      ["7/30/2016", "Beach", 82, 81, 0.35, 7744, 3249, 0],
      ["7/31/2016", "Beach", 82, 68, 0.35, 5776, 2209, 0],
    ]);

    const totalSalesRange = sheet.getRange("H2:H33");
    let data = [];
    for (let i = 2; i < 34; i++) {
      let item = [];
      item.push("=F" + i.toString() + "+G" + i.toString());
      data.push(item);
    }
    totalSalesRange.formulas = data;
    totalSalesRange.format.autofitColumns();

    const dateRange = sheet.getRange("A2:A33");
    const formatdate = [];
    const formatdateitem = ["m/d"];
    for (let i = 0; i < 32; i++) {
      formatdate.push(formatdateitem);
    }
    dateRange.numberFormat = formatdate;

    const numRange1 = sheet.getRange("F2:F33");
    const numRange2 = sheet.getRange("G2:G33");
    const numRange3 = sheet.getRange("H2:H33");
    let formatnumber = [];
    let formatnumberitem = ["###,0"];
    for (let i = 0; i < 32; i++) {
      formatnumber.push(formatnumberitem);
    }
    numRange1.numberFormat = formatnumber;
    numRange2.numberFormat = formatnumber;
    numRange3.numberFormat = formatnumber;

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
    sheet.activate();

    await context.sync();
    showStatus('Success for "Intialize sample data".', false);
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
