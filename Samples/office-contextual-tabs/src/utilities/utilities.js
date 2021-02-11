/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { getGlobal } from "../commands/commands.js";

/* global Excel, Office, console */

/**
 * Delete the sample worksheet with sales table
 */
async function deleteSampleWorkSheet() {
  return Excel.run(async context => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
  });
}

/**
 * Create a sample worksheet with sales data table. If the worksheet already exists,
 * replace it.
 * @param  {string} mockDataSource Identifies which mock data source to use to create the table.
 */
export async function createSampleWorkSheet(mockDataSource) {
  //Turn off the Refresh and Submit buttons.
  setSyncButtonEnabled(false);

  let g = getGlobal();
  g.isTableDirty = false; //reset dirty flag for new table

  //Ensure that the sample worksheet is deleted.
  await deleteSampleWorkSheet();
  Excel.run(async context => {
    //Create sample worksheet
    const sheet = context.workbook.worksheets.add("Sample");

    //Insert title row above table
    let range = sheet.getRange("A1");
    if (mockDataSource === "sqlMockData") {
      range.values = [["Data source: SQL Database"]];
    } else {
      range.values = [["Data source: External Excel File"]];
    }
    range.format.autofitColumns();

    //Add table with sales data
    let salesTable = sheet.tables.add("A2:E2", true);
    salesTable.name = "SalesTable";
    salesTable.onSelectionChanged.add(onSelectionChange);

    g.tableEventCount = 4; //This is to track and ignore the 4 events that will be generated from the next few lines of code.

    //add an onChanged event handler
    salesTable.onChanged.add(onChanged);

    //Add table header
    salesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];

    //Add data rows depending on which data source is in use.
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

    //creating is done
    g.isTableCreating = false;
    return context.sync();
  });
}

/**
 * Get the Sales table data and return as Promise in an array.
 */
export async function getTableData() {
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
  if (g.isTableSelected !== args.isInsideTable){
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

  //When the add-in creates the table, it will generate 4 events that we must ignore.
  //We only want to respond to the change events from the user.
  if (g.tableEventCount > 0) {
    g.tableEventCount--;
    return; //count down to throw away events caused by the table creation code
  }

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
export function setSyncButtonEnabled(visible) {
  let g = getGlobal();
  g.contextualTab.tabs[0].groups[1].controls[0].enabled = visible;
  g.contextualTab.tabs[0].groups[1].controls[1].enabled = visible;
  Office.ribbon.requestUpdate(g.contextualTab);
}
