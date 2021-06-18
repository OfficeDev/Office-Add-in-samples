/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, self, window, console */

 function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

/**
 * Handles all ribbon events from the Contoso contextual tab
 * @param  {} event: The event that was raised.
 */
function runRibbonAction(event){
  switch(event.source.id){
    case "btnSubmit": runSubmitAction();
        break;
    case "btnRefresh": runRefreshAction();
        break;
    case "itmExternalExcel": runImportExternalExcelFile();
        break;
    case "itmSQLSource": runImportSQLDatabase();
        break;
    default: console.log('Event ID: ' + event.source.id + ' was sent, but there is no function handler.');
    }
  event.completed();
}

/**
 * Submit data changes in table to data source
 */
function runSubmitAction(){
  const g = getGlobal();
  //Depending on which data source is in use, get data from the table, then update the mock data source.
  if (g.mockDataSource==='sqlMockData'){
    getTableData().then ((response) => {g.sqlMockData.data = response});
  } else if (g.mockDataSource==='excelFileMockData'){
    getTableData().then ((response) => {g.excelFileMockData.data = response}); 
  }
  //Turn off the Refresh and Submit buttons now that the table is in sync with data source.
  setSyncButtonEnabled(false);
  g.isTableDirty = false;
}

/**
 * Refresh the data in the table from the data source.
 */
function runRefreshAction(){
  //Recreate the table and sales data from source
  createSampleTable(g.mockDataSource);
  g.isTableDirty = false;
  setSyncButtonEnabled(false);
}

/**
 * Import data from mock External Excel file data source
 */
function runImportExternalExcelFile(){
  g.mockDataSource = 'excelFileMockData';
  //Just recreate the worksheet using the Excel file mock data source
  createSampleTable('excelFileMockData');
}

/**
 * Import data from mock SQL database source
 */
function runImportSQLDatabase(){
  g.mockDataSource = 'sqlMockData';
  //Just recreate the worksheet using the SQL database source
  createSampleTable(g.mockDataSource); 
}

// the add-in command functions need to be available in global scope
// Globals
const g = getGlobal();

let excelFileMockData = {data: [
  ["Frames", 5000, 7000, 6544, 4377],
  ["Saddles", 400, 323, 276, 651],
  ["Brake levers", 12000, 8766, 8456, 9812],
  ["Chains", 1550, 1088, 692, 853],
  ["Mirrors", 225, 600, 923, 544],
  ["Spokes", 6005, 7634, 4589, 8765]
]};

let sqlMockData = {data: [
  ["Frames", 10, 70, 654, 437],
  ["Saddles", 4000, 3230, 2760, 6510],
  ["Brake levers", 2000, 766, 456, 812],
  ["Chains", 150, 188, 62, 83],
  ["Mirrors", 25, 60, 93, 54],
  ["Spokes", 605, 734, 489, 875]
]};

g.contextualTab = getContextualRibbonJSON();
g.excelFileMockData = excelFileMockData;
g.sqlMockData = sqlMockData;
g.runRibbonAction = runRibbonAction;
g.mockDataSource = 'sqlMockData';
g.isTableSelected = false;
g.isTableDirty = false;