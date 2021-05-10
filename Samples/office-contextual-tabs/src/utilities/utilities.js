/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel, Office, console */

/**
 * Deletes the sample table.
 */
async function deleteSampleTable() {
  return Excel.run(async context => {
    try {
      const sheet = context.workbook.worksheets.getItem("Sample");
      let expensesTable = sheet.tables.getItem("SalesTable");
      expensesTable.delete();
      await context.sync();
    } catch (error) {
      if (error.code === "ItemNotFound") {
        return; //The table did not exist so just return.
      } else {
        console.log("deleteSampleTable failed");
        console.error(error);
        console.log(error.code);
        throw error; //Unexpected error occurred.
      }
    }
  });
}

/**
 * Create the sales data table. If the table already exists, replace it.
 * @param  {string} mockDataSource Identifies which mock data source to use to create the table.
 */
 async function createSampleTable(mockDataSource) {
  //Delete table if it already exists
  await deleteSampleTable();
  console.log("creating table");
  return Excel.run(async context => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    //Create title row above table
    let range = sheet.getRange("A1");
    if (mockDataSource === "sqlMockData") {
      range.values = [["Data source: SQL Database"]];
    } else {
      range.values = [["Data source: External Excel File"]];
    }
    range.format.autofitColumns();

    //Create table
    let salesTable = sheet.tables.add("A2:E2", true);
    salesTable.name = "SalesTable";

    //Add table header
    salesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];

    //Add data rows depending on which data source is in use.
    const g = getGlobal();
    if (mockDataSource === "sqlMockData") {
      salesTable.rows.add(null, g.sqlMockData.data);
    } else if (mockDataSource === "excelFileMockData") {
      salesTable.rows.add(null, g.excelFileMockData.data);
    }

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
    sheet.activate();
    sheet.getRange("A2").select();
    await context.sync();

    //Add event handlers
    salesTable.onSelectionChanged.add(onSelectionChange);
    salesTable.onChanged.add(onChanged);

    return context.sync();
  });
}

/**
 * Create the sample worksheet.
 */
 async function createSampleWorkSheet() {
  Excel.run(async context => {
    try {
      //Create sample worksheet
      context.workbook.worksheets.add("Sample");
      await context.sync();
    } catch (error) {
      if (error.code === "ItemAlreadyExists") {
        return; //The worksheet already exists so just return.
      } else {
        console.error(error);
        throw error; //Unexpected error occurred.
      }
    }
  });
  return;
}

/**
 * Get the Sales table data and return as Promise in an array.
 */
 async function getTableData() {
  let response = null;

  return Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const expensesTable = sheet.tables.getItem("SalesTable");
    const bodyRange = expensesTable.getDataBodyRange().load("values");
    await context.sync();

    response = bodyRange.values;
    return response;
  });
}

/**
 * Handles the onSelectionChange event. If selection is inside the table, the Contoso custom tab is shown.
 * Otherwise the Contoso custom tab is hidden.
 * @param  {} args The arguments for the selection changed event.
 */
function onSelectionChange(args) {
  let g = getGlobal();
  if (g.isTableSelected !== args.isInsideTable) {
    g.isTableSelected = args.isInsideTable;
    setContextualTabVisibility(args.isInsideTable);
  }
}

/**
 * Handles the onChanged event. When data in the sales table is changed,
 * enable the refresh and submit buttons.
 */
function onChanged() {
  let g = getGlobal();
  //check if dirty flag was set (flag avoids extra unnecessary ribbon operations)
  if (!g.isTableDirty) {
    g.isTableDirty = true;

    //Enable the Refresh and Submit buttons
    setSyncButtonEnabled(true);
  }
}

/**
 * Shows or hides the contextual tab for Contoso depending on the visible parameter.
 * @param  {} visible true if the contextual tab is to be shown; otherwise, false.
 */
function setContextualTabVisibility(visible) {
  let g = getGlobal();
  g.contextualTab.tabs[0].visible = visible;
  try {
    Office.ribbon.requestUpdate(g.contextualTab);
  } catch (err) {
    console.error(err);
  }
}

/**
 * Enables or disables the Refresh and Submit buttons for table data.
 *
 * @param  {boolean} visible true if the buttons should be enabled; otherwise, false.
 */
 function setSyncButtonEnabled(visible) {
  console.log("sync is " + visible);
  let g = getGlobal();
  g.contextualTab.tabs[0].groups[1].controls[0].enabled = visible;
  g.contextualTab.tabs[0].groups[1].controls[1].enabled = visible;
  console.log(g.contextualTab);
  Office.ribbon.requestUpdate(g.contextualTab);
}
